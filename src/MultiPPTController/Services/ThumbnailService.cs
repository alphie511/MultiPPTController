using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Docnet.Core;
using Docnet.Core.Converters;
using Docnet.Core.Models;

namespace MultiPPTController.Services;

/// <summary>
/// 从演示文稿提取封面缩略图：PPT 读内嵌预览或 WPS 导出，PDF 渲染首页。
/// </summary>
public sealed class ThumbnailService
{
    private static readonly object PdfRenderLock = new();

    private readonly string _cacheDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "MultiPPTController",
        "thumbs");

    /// <summary>
    /// 生成或读取缓存后的封面图。
    /// </summary>
    /// <param name="pptPath">演示文稿路径。</param>
    /// <returns>可绑定的图像；失败则为 null。</returns>
    public ImageSource? Create(string pptPath)
    {
        try
        {
            Directory.CreateDirectory(_cacheDir);
            var dest = Path.Combine(_cacheDir, HashKey(pptPath) + ".png");
            if (!File.Exists(dest))
            {
                var created = DeckConvertService.IsPdf(pptPath)
                    ? TryRenderPdfPage(pptPath, dest)
                    : TryExtractEmbedded(pptPath, dest) || ExportViaWps(pptPath, dest);
                if (!created)
                {
                    return null;
                }
            }

            return LoadBitmap(dest);
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// 按文件路径与修改时间生成缓存键。
    /// </summary>
    /// <param name="pptPath">文件路径。</param>
    /// <returns>十六进制哈希。</returns>
    private static string HashKey(string pptPath)
    {
        var stamp = File.GetLastWriteTimeUtc(pptPath).Ticks;
        var bytes = Encoding.UTF8.GetBytes(pptPath + "|" + stamp);
        return Convert.ToHexString(SHA256.HashData(bytes))[..16];
    }

    /// <summary>
    /// 用 PDFium 渲染 PDF 第一页为 PNG。
    /// </summary>
    /// <param name="pdfPath">PDF 路径。</param>
    /// <param name="dest">输出 PNG 路径。</param>
    /// <returns>成功则为 true。</returns>
    private static bool TryRenderPdfPage(string pdfPath, string dest)
    {
        lock (PdfRenderLock)
        {
            try
            {
                using var reader = DocLib.Instance.GetDocReader(pdfPath, new PageDimensions(360, 640));
                if (reader.GetPageCount() < 1)
                {
                    return false;
                }

                using var page = reader.GetPageReader(0);
                var pixels = page.GetImage(new NaiveTransparencyRemover(255, 255, 255));
                var width = page.GetPageWidth();
                var height = page.GetPageHeight();
                if (pixels.Length < width * height * 4)
                {
                    return false;
                }

                var bitmap = BitmapSource.Create(
                    width,
                    height,
                    96,
                    96,
                    PixelFormats.Bgra32,
                    null,
                    pixels,
                    width * 4);
                bitmap.Freeze();
                var encoder = new PngBitmapEncoder();
                encoder.Frames.Add(BitmapFrame.Create(bitmap));
                using var stream = File.Create(dest);
                encoder.Save(stream);
                return File.Exists(dest) && new FileInfo(dest).Length > 0;
            }
            catch
            {
                try
                {
                    if (File.Exists(dest))
                    {
                        File.Delete(dest);
                    }
                }
                catch
                {
                    // 清理失败不影响返回
                }

                return false;
            }
        }
    }

    /// <summary>
    /// 尝试从 Office 压缩包取出内嵌缩略图。
    /// </summary>
    /// <param name="pptPath">演示文稿路径。</param>
    /// <param name="dest">输出 PNG/JPEG 路径。</param>
    /// <returns>成功则为 true。</returns>
    private static bool TryExtractEmbedded(string pptPath, string dest)
    {
        try
        {
            using var zip = ZipFile.OpenRead(pptPath);
            foreach (var name in new[]
                     {
                         "docProps/thumbnail.jpeg",
                         "docProps/thumbnail.jpg",
                         "docProps/thumbnail.png"
                     })
            {
                var entry = zip.GetEntry(name);
                if (entry is null)
                {
                    continue;
                }

                entry.ExtractToFile(dest, overwrite: true);
                return File.Exists(dest) && new FileInfo(dest).Length > 0;
            }
        }
        catch
        {
            // 旧版 .ppt / 非 zip 容器
        }

        return false;
    }

    /// <summary>
    /// 用 WPS COM 导出第一页为 PNG。
    /// </summary>
    /// <param name="pptPath">演示文稿路径。</param>
    /// <param name="dest">输出路径。</param>
    /// <returns>成功则为 true。</returns>
    private static bool ExportViaWps(string pptPath, string dest)
    {
        if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
        {
            string? error = null;
            var ok = false;
            var thread = new Thread(() =>
            {
                try
                {
                    ok = ExportViaWps(pptPath, dest);
                }
                catch (Exception ex)
                {
                    error = ex.Message;
                }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();
            return ok && error is null;
        }

        var type = Type.GetTypeFromProgID("KWPP.Application", throwOnError: false);
        if (type is null)
        {
            return false;
        }

        dynamic? app = null;
        dynamic? presentation = null;
        try
        {
            app = Activator.CreateInstance(type);
            try
            {
                app!.Visible = false;
            }
            catch
            {
                // 个别版本必须可见
            }

            presentation = app?.Presentations.Open(pptPath, -1, 0, 0);
            if (presentation is null)
            {
                return false;
            }

            presentation.Slides.Item(1).Export(dest, "PNG", 640, 360);
            return File.Exists(dest);
        }
        catch
        {
            return false;
        }
        finally
        {
            try
            {
                presentation?.Close();
            }
            catch
            {
                // 忽略
            }

            try
            {
                app?.Quit();
            }
            catch
            {
                // 忽略
            }
        }
    }

    /// <summary>
    /// 从磁盘加载并冻结位图，避免锁定文件。
    /// </summary>
    /// <param name="path">图像路径。</param>
    /// <returns>位图源。</returns>
    private static ImageSource LoadBitmap(string path)
    {
        var image = new BitmapImage();
        image.BeginInit();
        image.CacheOption = BitmapCacheOption.OnLoad;
        image.UriSource = new Uri(path);
        image.EndInit();
        image.Freeze();
        return image;
    }
}
