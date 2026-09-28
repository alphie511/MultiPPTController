using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Threading;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.Wpf;
using MultiPPTController.Models;
using MultiPPTController.Services;

namespace MultiPPTController;

/// <summary>
/// 无边框 WebView2 放映窗：PPTX 走 vanilla viewer，PDF 走 pdf.js；可铺主屏或副屏。
/// </summary>
public sealed class WebPlaybackWindow : Window
{
    public const string WindowTitle = "MultiPPT-web";

    private readonly WebView2 _webView = new();
    private static CoreWebView2Environment? _sharedEnvironment;
    private bool _ready;
    private int _knownSlide = 1;

    /// <summary>主屏播放条上一页/下一页，交给编排器同步副屏。</summary>
    public event Action<string>? HostCommand;

    /// <summary>
    /// 创建铺满目标屏的黑色无边框宿主。
    /// </summary>
    /// <param name="screen">主屏。</param>
    public WebPlaybackWindow(DisplayScreen screen)
    {
        Screen = screen;
        Title = WindowTitle;
        WindowStyle = WindowStyle.None;
        ResizeMode = ResizeMode.NoResize;
        ShowInTaskbar = false;
        Topmost = true;
        WindowStartupLocation = WindowStartupLocation.Manual;
        Background = System.Windows.Media.Brushes.Black;
        Content = _webView;
    }

    /// <summary>目标屏幕。</summary>
    public DisplayScreen Screen { get; }

    /// <summary>Win32 句柄；尚未创建则为 Zero。</summary>
    public IntPtr Hwnd { get; private set; }

    /// <summary>
    /// 初始化 WebView2、映射本地 PlayerHost，并打开文稿。
    /// </summary>
    /// <param name="deckPath">可被浏览器打开的 .pptx / .pdf 路径。</param>
    /// <returns>完成任务。</returns>
    public async Task StartAsync(string deckPath)
    {
        Show();
        Hwnd = new WindowInteropHelper(this).EnsureHandle();
        NativeWindowHelper.FitToBounds(Hwnd, Screen.Bounds, raiseZOrder: true);

        var environment = await GetSharedEnvironmentAsync().ConfigureAwait(true);
        await _webView.EnsureCoreWebView2Async(environment).ConfigureAwait(true);

        var hostDir = ResolvePlayerHostDirectory();
        _webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
        _webView.CoreWebView2.Settings.AreDevToolsEnabled = false;
        _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;
        _webView.CoreWebView2.SetVirtualHostNameToFolderMapping(
            "player.multippt",
            hostDir,
            CoreWebView2HostResourceAccessKind.Allow);

        var runtimeDir = Path.Combine(hostDir, "runtime");
        Directory.CreateDirectory(runtimeDir);
        var runtimeName = $"{Guid.NewGuid():N}{Path.GetExtension(deckPath)}";
        var runtimePath = Path.Combine(runtimeDir, runtimeName);
        File.Copy(deckPath, runtimePath, overwrite: true);

        var loaded = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        void OnNavigation(object? sender, CoreWebView2NavigationCompletedEventArgs args)
        {
            _webView.CoreWebView2.NavigationCompleted -= OnNavigation;
            if (args.IsSuccess)
            {
                loaded.TrySetResult();
            }
            else
            {
                loaded.TrySetException(new InvalidOperationException($"Web 播放器页加载失败：{args.WebErrorStatus}"));
            }
        }

        _webView.CoreWebView2.NavigationCompleted += OnNavigation;
        _webView.CoreWebView2.Navigate($"https://player.multippt/index.html?t={DateTime.UtcNow.Ticks}");
        await loaded.Task.WaitAsync(TimeSpan.FromSeconds(20)).ConfigureAwait(true);
        NativeWindowHelper.FitToBounds(Hwnd, Screen.Bounds, raiseZOrder: true);

        var url = $"https://player.multippt/runtime/{runtimeName}";
        var script = $"window.player.load({JsonSerializer.Serialize(url)})";
        await _webView.ExecuteScriptAsync(script).WaitAsync(TimeSpan.FromSeconds(30)).ConfigureAwait(true);
        _ready = true;
        NativeWindowHelper.FitToBounds(Hwnd, Screen.Bounds, raiseZOrder: true);
    }

    /// <summary>
    /// 调用页面里的 player API（可在 Web STA 线程同步等待，不卡死消息泵）。
    /// </summary>
    /// <param name="expression">JS 表达式，应返回 JSON 可序列化对象。</param>
    /// <returns>解析后的页码状态。</returns>
    public PlaybackProgress InvokePlayer(string expression)
    {
        if (!_ready)
        {
            return default;
        }

        var task = _webView.ExecuteScriptAsync($"(() => {{ const value = {expression}; return value; }})()");
        if (!task.IsCompleted)
        {
            var frame = new DispatcherFrame();
            task.ContinueWith(static (_, state) =>
            {
                if (state is DispatcherFrame dispatcherFrame)
                {
                    dispatcherFrame.Continue = false;
                }
            }, frame);
            Dispatcher.PushFrame(frame);
        }

        var parsed = ParseState(task.GetAwaiter().GetResult());
        if (parsed.HasSlide)
        {
            _knownSlide = parsed.SlideIndex;
            return parsed;
        }

        return new PlaybackProgress(_knownSlide, 0);
    }

    /// <summary>
    /// 接收页面里播放条/单击发出的 next、prev、stop。
    /// </summary>
    /// <param name="sender">WebView2。</param>
    /// <param name="args">消息。</param>
    private void OnWebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs args)
    {
        try
        {
            using var document = JsonDocument.Parse(args.WebMessageAsJson);
            if (document.RootElement.TryGetProperty("cmd", out var command) &&
                command.GetString() is { Length: > 0 } text)
            {
                HostCommand?.Invoke(text);
            }
        }
        catch
        {
            // 非法消息忽略
        }
    }

    /// <summary>
    /// 按目标屏重新铺满本窗。
    /// </summary>
    public void Relayout()
    {
        if (Hwnd != IntPtr.Zero)
        {
            NativeWindowHelper.FitToBounds(Hwnd, Screen.Bounds, raiseZOrder: false);
        }
    }

    /// <summary>
    /// 关闭 WebView2 与窗口。
    /// </summary>
    public void Shutdown()
    {
        try
        {
            _webView.Dispose();
        }
        catch
        {
            // 退出阶段忽略
        }

        Close();
    }

    /// <summary>
    /// 多路 Web 窗共用同一 WebView2 用户目录，避免副屏 PDF 再开环境失败。
    /// </summary>
    /// <returns>共享环境。</returns>
    private static async Task<CoreWebView2Environment> GetSharedEnvironmentAsync()
    {
        if (_sharedEnvironment is not null)
        {
            return _sharedEnvironment;
        }

        var userData = Path.Combine(Path.GetTempPath(), "MultiPPTController", "webview2");
        Directory.CreateDirectory(userData);
        _sharedEnvironment = await CoreWebView2Environment.CreateAsync(null, userData).ConfigureAwait(true);
        return _sharedEnvironment;
    }

    /// <summary>
    /// 定位输出目录中的 PlayerHost。
    /// </summary>
    /// <returns>本地目录。</returns>
    private static string ResolvePlayerHostDirectory()
    {
        var dir = Path.Combine(AppContext.BaseDirectory, "PlayerHost");
        if (!File.Exists(Path.Combine(dir, "index.html")))
        {
            throw new InvalidOperationException("未找到 PlayerHost/index.html，请先构建主屏播放器资源。");
        }

        if (!File.Exists(Path.Combine(dir, "vendor", "player.js")))
        {
            throw new InvalidOperationException("未找到 PlayerHost/vendor/player.js。");
        }

        return dir;
    }

    /// <summary>
    /// 解析 WebView2 ExecuteScriptAsync 返回的 JSON 字符串。
    /// </summary>
    /// <param name="raw">双编码 JSON。</param>
    /// <returns>页进度。</returns>
    private static PlaybackProgress ParseState(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw) || raw == "null")
        {
            return default;
        }

        try
        {
            var json = raw.Trim();
            if (json.Length >= 2 && json[0] == '"')
            {
                json = JsonSerializer.Deserialize<string>(json) ?? raw;
            }

            using var document = JsonDocument.Parse(json);
            var root = document.RootElement;
            var slide = root.TryGetProperty("slide", out var slideNode) ? slideNode.GetInt32() : 0;
            var click = root.TryGetProperty("click", out var clickNode) ? clickNode.GetInt32() : 0;
            return new PlaybackProgress(slide, click);
        }
        catch
        {
            return default;
        }
    }
}
