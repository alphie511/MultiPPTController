namespace MultiPPTController.Services;

/// <summary>
/// 把 WPS 专有格式转成 Web 播放器能打开的临时 .pptx。
/// </summary>
public static class DeckConvertService
{
    private const int MsoTrue = -1;
    private const int MsoFalse = 0;
    private const int PpSaveAsOpenXmlPresentation = 24;

    /// <summary>
    /// 浏览器原生可打开的扩展名。
    /// </summary>
    private static readonly HashSet<string> WebNativeExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".pptx", ".pptm", ".ppsx", ".ppt", ".potx"
    };

    /// <summary>
    /// 是否可直接交给 Web 播放器，无需 WPS 另存。
    /// </summary>
    /// <param name="path">文稿路径。</param>
    /// <returns>可直接打开则为 true。</returns>
    public static bool IsWebNative(string path)
    {
        return WebNativeExtensions.Contains(Path.GetExtension(path));
    }

    /// <summary>
    /// 是否为 PDF（主屏与副屏都走 WebView2 + pdf.js）。
    /// </summary>
    /// <param name="path">文稿路径。</param>
    /// <returns>PDF 则为 true。</returns>
    public static bool IsPdf(string path)
    {
        return string.Equals(Path.GetExtension(path), ".pdf", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// 主屏 Web 用的打开路径：pptx 原样返回，.dps/.dpt 经 WPS 另存临时 pptx。
    /// </summary>
    /// <param name="application">已启动的 KWPP.Application；无需转换时可为 null。</param>
    /// <param name="sourcePath">原始文稿。</param>
    /// <returns>Web 可打开的路径；ownsTemp 表示结束时需删除。</returns>
    public static (string Path, bool OwnsTemp) EnsureWebPlayable(dynamic? application, string sourcePath)
    {
        if (IsWebNative(sourcePath))
        {
            return (sourcePath, false);
        }

        if (application is null)
        {
            throw new InvalidOperationException($"主屏 Web 播放器无法直接打开 {Path.GetExtension(sourcePath)}，需要 WPS 先另存为 .pptx。");
        }

        var dir = Path.Combine(Path.GetTempPath(), "MultiPPTController", "web");
        Directory.CreateDirectory(dir);
        var dest = Path.Combine(dir, $"{Guid.NewGuid():N}.pptx");
        var presentation = application.Presentations.Open(sourcePath, MsoTrue, MsoFalse, MsoFalse);
        try
        {
            presentation.SaveAs(dest, PpSaveAsOpenXmlPresentation);
        }
        finally
        {
            try
            {
                presentation.Close();
            }
            catch
            {
                // 忽略重复关闭
            }
        }

        return (dest, true);
    }
}
