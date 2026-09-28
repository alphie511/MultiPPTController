using MultiPPTController.Models;

namespace MultiPPTController.Services;

/// <summary>
/// WebView2 放映会话：PPTX 用 vanilla viewer，PDF 用 pdf.js。
/// </summary>
public sealed class WebPlaybackSession : IScreenPlayer
{
    private readonly WebPlayerPump _pump;
    private readonly bool _ownsTempCopy;
    private WebPlaybackWindow? _window;

    /// <summary>
    /// 绑定片库条目、主屏与打开路径。
    /// </summary>
    /// <param name="deck">片库条目。</param>
    /// <param name="screen">主屏。</param>
    /// <param name="pump">WebView2 STA 泵。</param>
    /// <param name="openPath">实际打开的 .pptx 或 .pdf 路径。</param>
    /// <param name="ownsTempCopy">结束时是否删除临时副本。</param>
    public WebPlaybackSession(
        DeckItem deck,
        DisplayScreen screen,
        WebPlayerPump pump,
        string openPath,
        bool ownsTempCopy)
    {
        Deck = deck;
        Screen = screen;
        OpenPath = openPath;
        _pump = pump;
        _ownsTempCopy = ownsTempCopy;
    }

    /// <inheritdoc />
    public DeckItem Deck { get; }

    /// <inheritdoc />
    public DisplayScreen Screen { get; }

    /// <summary>Web 打开用的文件路径。</summary>
    public string OpenPath { get; }

    /// <inheritdoc />
    public IntPtr ShowHwnd => _window?.Hwnd ?? IntPtr.Zero;

    /// <summary>主屏播放条发出的 next / prev / stop。</summary>
    public event Action<string>? HostCommand;

    /// <summary>
    /// 创建无边框窗并加载文稿。
    /// </summary>
    /// <returns>完成任务。</returns>
    public Task StartAsync()
    {
        return _pump.InvokeAsync(async () =>
        {
            _window = new WebPlaybackWindow(Screen);
            _window.HostCommand += command => HostCommand?.Invoke(command);
            await _window.StartAsync(OpenPath).ConfigureAwait(true);
            try
            {
                var progress = _window.InvokePlayer("window.player.state()");
                if (progress.HasSlide)
                {
                    Deck.SlideCount = Math.Max(Deck.SlideCount ?? 0, progress.SlideIndex);
                }
            }
            catch
            {
                // 页数读失败不影响开播
            }
        });
    }

    /// <inheritdoc />
    public string DescribeShow()
    {
        var hwnd = ShowHwnd;
        var progress = ReadProgress();
        return $"{Deck.DisplayName} @{Screen.Label} web=0x{hwnd.ToInt64():X} '{NativeWindowHelper.GetTitle(hwnd)}' {progress}";
    }

    /// <inheritdoc />
    public PlaybackProgress ReadProgress()
    {
        if (_window is null)
        {
            return default;
        }

        try
        {
            return _pump.Invoke(() => _window.InvokePlayer("window.player.state()"));
        }
        catch
        {
            return default;
        }
    }

    /// <inheritdoc />
    public void Next()
    {
        if (_window is null)
        {
            return;
        }

        _pump.Invoke(() => _window.InvokePlayer("window.player.next()"));
    }

    /// <inheritdoc />
    public void Previous()
    {
        if (_window is null)
        {
            return;
        }

        _pump.Invoke(() => _window.InvokePlayer("window.player.prev()"));
    }

    /// <inheritdoc />
    public void RelayoutToScreen(IEnumerable<IntPtr>? occupiedByOthers = null, bool raiseZOrder = false)
    {
        if (_window is null)
        {
            return;
        }

        _pump.Invoke(() => _window.Relayout());
    }

    /// <inheritdoc />
    public void WarmUp()
    {
        RelayoutToScreen(raiseZOrder: true);
        try
        {
            _pump.Invoke(() => _window?.InvokePlayer("window.player.state()"));
        }
        catch
        {
            // 预热失败不影响后续翻页
        }
    }

    /// <inheritdoc />
    public void Stop()
    {
        if (_window is not null)
        {
            try
            {
                _pump.Invoke(() =>
                {
                    _window?.Shutdown();
                    _window = null;
                });
            }
            catch
            {
                _window = null;
            }
        }

        Deck.Status = "已停止";
        if (!_ownsTempCopy)
        {
            return;
        }

        try
        {
            if (File.Exists(OpenPath))
            {
                File.Delete(OpenPath);
            }
        }
        catch
        {
            // 临时副本删除失败不影响结束放映
        }
    }
}
