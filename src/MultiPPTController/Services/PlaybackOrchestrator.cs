using MultiPPTController.Models;

namespace MultiPPTController.Services;

/// <summary>
/// 按文件与屏幕分流：PPTX 主屏 Web / 副屏 WPS；PDF 任意屏都走 WebView2。
/// </summary>
public sealed class PlaybackOrchestrator : IDisposable
{
    private readonly StaComPump _pump = new();
    private readonly WebPlayerPump _webPump = new();
    private readonly List<IScreenPlayer> _players = [];
    private WpsEngineHost? _wpsHost;
    private CancellationTokenSource? _layoutWatch;
    private bool _disposed;

    /// <summary>当前放映阶段。</summary>
    public PlaybackStatus Status { get; private set; } = PlaybackStatus.Idle;

    /// <summary>
    /// 探测本机是否已注册 WPS 演示 COM。
    /// </summary>
    /// <returns>可创建则为 true。</returns>
    public static bool IsWpsAvailable()
    {
        return Type.GetTypeFromProgID("KWPP.Application", throwOnError: false) is not null;
    }

    /// <summary>
    /// 打开全部文稿：PPTX 主屏 Web / 副屏 WPS；PDF 任意屏 Web。
    /// </summary>
    /// <param name="assignments">文稿与屏幕配对。</param>
    /// <returns>完成任务。</returns>
    public async Task StartAsync(IReadOnlyList<(DeckItem Deck, DisplayScreen Screen)> assignments)
    {
        Status = PlaybackStatus.Starting;
        try
        {
            var webAssignments = assignments
                .Where(item => item.Screen.IsPrimary || DeckConvertService.IsPdf(item.Deck.FilePath))
                .ToList();
            var wpsAssignments = assignments
                .Where(item => !item.Screen.IsPrimary && !DeckConvertService.IsPdf(item.Deck.FilePath))
                .ToList();
            var needWps = wpsAssignments.Count > 0 ||
                          webAssignments.Any(item =>
                              !DeckConvertService.IsPdf(item.Deck.FilePath) &&
                              !DeckConvertService.IsWebNative(item.Deck.FilePath));

            if (needWps)
            {
                await _pump.InvokeAsync(() =>
                {
                    _wpsHost = WpsEngineHost.Launch();
                }).ConfigureAwait(false);
            }

            foreach (var (deck, screen) in webAssignments)
            {
                var (path, ownsTemp) = await ResolveWebPathAsync(deck.FilePath).ConfigureAwait(false);
                var session = new WebPlaybackSession(deck, screen, _webPump, path, ownsTemp);
                session.HostCommand += OnWebHostCommand;
                await session.StartAsync().ConfigureAwait(false);
                _players.Add(session);
            }

            if (wpsAssignments.Count > 0)
            {
                await _pump.InvokeAsync(() =>
                {
                    _wpsHost ??= WpsEngineHost.Launch();
                    var openedPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    foreach (var (deck, screen) in wpsAssignments)
                    {
                        var ownsTemp = false;
                        var path = deck.FilePath;
                        if (!openedPaths.Add(path))
                        {
                            path = ClonePresentation(deck.FilePath);
                            ownsTemp = true;
                        }

                        _wpsHost.StartSession(deck, screen, path, ownsTemp);
                    }

                    foreach (var session in _wpsHost.Sessions)
                    {
                        _players.Add(session);
                    }

                    NativeWindowHelper.HideEmptyEngineOverlays();
                }).ConfigureAwait(false);
            }
            else if (_wpsHost is not null)
            {
                await _pump.InvokeAsync(() =>
                {
                    _wpsHost.Quit();
                    _wpsHost = null;
                }).ConfigureAwait(false);
            }

            ApplyPlayingStatus();
            await RelayoutAllAsync().ConfigureAwait(false);
            foreach (var session in _players)
            {
                session.WarmUp();
            }

            Status = PlaybackStatus.Playing;
        }
        catch
        {
            await TearDownAsync().ConfigureAwait(false);
            Status = PlaybackStatus.Idle;
            throw;
        }
    }

    /// <summary>
    /// 各屏按进度对齐后前进一步：已超前的屏等待，只给落后的屏发下一步。
    /// </summary>
    /// <returns>给界面看的同步说明；已对齐则为 null。</returns>
    public Task<string?> NextAsync()
    {
        return AdvanceAlignedAsync(forward: true);
    }

    /// <summary>
    /// 各屏按进度对齐后后退一步：只让超前的屏回退。
    /// </summary>
    /// <returns>给界面看的同步说明；已对齐则为 null。</returns>
    public Task<string?> PreviousAsync()
    {
        return AdvanceAlignedAsync(forward: false);
    }

    /// <summary>
    /// 诊断：各路已绑句柄与放映状态。
    /// </summary>
    /// <returns>每路一行。</returns>
    public Task<IReadOnlyList<string>> DescribeSessionsAsync()
    {
        return Task.FromResult<IReadOnlyList<string>>(_players.Select(session => session.DescribeShow()).ToList());
    }

    /// <summary>
    /// 读取各路当前页码，供冒烟核对连续翻页。
    /// </summary>
    /// <returns>每路一行「标签=进度」。</returns>
    public Task<IReadOnlyList<string>> ReadProgressLinesAsync()
    {
        return _pump.InvokeAsync<IReadOnlyList<string>>(() =>
            _players.Select(session => $"{session.Screen.Label}\t{session.ReadProgress()}").ToList());
    }

    /// <summary>
    /// 主界面收起后再按各屏物理分辨率铺一次。
    /// </summary>
    /// <returns>完成任务。</returns>
    public async Task RelayoutAllAsync()
    {
        var occupied = _players.Select(session => session.ShowHwnd).ToList();
        foreach (var session in _players.OfType<WebPlaybackSession>())
        {
            if (NativeWindowHelper.MatchesBounds(session.ShowHwnd, session.Screen.Bounds))
            {
                continue;
            }

            session.RelayoutToScreen(occupied.Where(hwnd => hwnd != session.ShowHwnd));
        }

        if (_players.OfType<WpsPlaybackSession>().Any())
        {
            await _pump.InvokeAsync(() =>
            {
                var live = _players.Select(session => session.ShowHwnd).ToList();
                foreach (var session in _players.OfType<WpsPlaybackSession>())
                {
                    if (NativeWindowHelper.MatchesBounds(session.ShowHwnd, session.Screen.Bounds, 24))
                    {
                        continue;
                    }

                    session.RelayoutToScreen(live.Where(hwnd => hwnd != session.ShowHwnd));
                }
            }).ConfigureAwait(false);
        }
    }

    /// <summary>
    /// 结束全部放映并关闭已打开的文稿。
    /// </summary>
    /// <returns>完成任务。</returns>
    public async Task StopAsync()
    {
        Status = PlaybackStatus.Stopping;
        await TearDownAsync().ConfigureAwait(false);
        Status = PlaybackStatus.Idle;
    }

    /// <summary>
    /// 释放 COM、Web 泵与 STA 线程。
    /// </summary>
    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;
        StopLayoutWatch();
        try
        {
            TearDownAsync().GetAwaiter().GetResult();
        }
        catch
        {
            // 退出阶段忽略
        }

        _pump.Dispose();
        _webPump.Dispose();
    }

    /// <summary>
    /// 主屏 Web 打开路径：pptx / pdf 原样用；.dps/.dpt 经 WPS 另存临时 pptx。
    /// </summary>
    /// <param name="sourcePath">原始文稿。</param>
    /// <returns>打开路径与是否为临时文件。</returns>
    private Task<(string Path, bool OwnsTemp)> ResolveWebPathAsync(string sourcePath)
    {
        if (DeckConvertService.IsPdf(sourcePath) || DeckConvertService.IsWebNative(sourcePath))
        {
            return Task.FromResult((sourcePath, false));
        }

        return _pump.InvokeAsync<(string Path, bool OwnsTemp)>(() =>
        {
            _wpsHost ??= WpsEngineHost.Launch();
            var converted = DeckConvertService.EnsureWebPlayable(_wpsHost.Application, sourcePath);
            return ((string)converted.Path, (bool)converted.OwnsTemp);
        });
    }

    /// <summary>
    /// 同一 PPT 多路 WPS 时复制一份临时文件。
    /// </summary>
    /// <param name="sourcePath">原始文稿路径。</param>
    /// <returns>临时副本路径。</returns>
    private static string ClonePresentation(string sourcePath)
    {
        var dir = Path.Combine(Path.GetTempPath(), "MultiPPTController", "play");
        Directory.CreateDirectory(dir);
        var dest = Path.Combine(dir, $"{Guid.NewGuid():N}{Path.GetExtension(sourcePath)}");
        File.Copy(sourcePath, dest, overwrite: true);
        return dest;
    }

    /// <summary>
    /// 放映中持续校正副屏分辨率。
    /// </summary>
    private void StartLayoutWatch()
    {
        StopLayoutWatch();
        var watch = new CancellationTokenSource();
        _layoutWatch = watch;
        var token = watch.Token;
        _ = Task.Run(async () =>
        {
            var elapsed = 0;
            while (!token.IsCancellationRequested && elapsed < 2000)
            {
                var delay = 400;
                try
                {
                    await Task.Delay(delay, token).ConfigureAwait(false);
                }
                catch (OperationCanceledException)
                {
                    return;
                }

                elapsed += delay;
                if (Status != PlaybackStatus.Playing)
                {
                    return;
                }

                try
                {
                    await RelayoutAllAsync().ConfigureAwait(false);
                }
                catch
                {
                    // 已停播时忽略
                }
            }
        }, token);
    }

    /// <summary>
    /// 停止分辨率校正循环。
    /// </summary>
    private void StopLayoutWatch()
    {
        if (_layoutWatch is null)
        {
            return;
        }

        _layoutWatch.Cancel();
        _layoutWatch.Dispose();
        _layoutWatch = null;
    }

    /// <summary>
    /// 按每份文稿实际占用的屏幕刷新「放映中」状态。
    /// </summary>
    private void ApplyPlayingStatus()
    {
        foreach (var group in _players.GroupBy(session => session.Deck))
        {
            var labels = group.Select(session => session.Screen.Label).ToList();
            group.Key.Status = labels.Count == 1
                ? $"放映中 · {labels[0]}"
                : $"放映中 · {string.Join("、", labels)}";
        }
    }

    /// <summary>
    /// 混合引擎时各走一步（Web 与 WPS 页码不可比）；纯 WPS 才做超前等待。
    /// </summary>
    /// <param name="forward">true 为下一步，false 为上一步。</param>
    /// <returns>需要提示用户的说明；已对齐则为 null。</returns>
    private Task<string?> AdvanceAlignedAsync(bool forward)
    {
        return _pump.InvokeAsync<string?>(() =>
        {
            if (_players.Any(player => player is WebPlaybackSession))
            {
                foreach (var session in _players)
                {
                    TryStep(session, forward);
                }

                return null;
            }

            var snapshots = _players
                .Select(session =>
                {
                    PlaybackProgress progress;
                    try
                    {
                        progress = session.ReadProgress();
                    }
                    catch
                    {
                        progress = default;
                    }

                    return (Session: session, Progress: progress);
                })
                .ToList();

            if (snapshots.Count == 0)
            {
                return null;
            }

            var readable = snapshots.Where(item => item.Progress.HasSlide).ToList();
            if (readable.Count == 0)
            {
                foreach (var item in snapshots)
                {
                    TryStep(item.Session, forward);
                }

                return null;
            }

            if (forward)
            {
                var behind = readable.Min(item => item.Progress);
                var aligned = readable.All(item => item.Progress.CompareTo(behind) == 0);
                foreach (var item in snapshots)
                {
                    if (aligned || !item.Progress.HasSlide || item.Progress.CompareTo(behind) == 0)
                    {
                        TryStep(item.Session, forward: true);
                    }
                }

                return aligned ? null : BuildWaitMessage(readable, behind, waitingAhead: true);
            }

            var ahead = readable.Max(item => item.Progress);
            var allSame = readable.All(item => item.Progress.CompareTo(ahead) == 0);
            foreach (var item in snapshots)
            {
                if (allSame || !item.Progress.HasSlide || item.Progress.CompareTo(ahead) == 0)
                {
                    TryStep(item.Session, forward: false);
                }
            }

            return allSame ? null : BuildWaitMessage(readable, ahead, waitingAhead: false);
        });
    }

    /// <summary>
    /// 主屏播放条/单击翻页时，两边一起走，从而带动副屏。
    /// </summary>
    /// <param name="command">next、prev 或 stop。</param>
    private void OnWebHostCommand(string command)
    {
        if (Status != PlaybackStatus.Playing)
        {
            return;
        }

        _ = HandleWebHostCommandAsync(command);
    }

    /// <summary>
    /// 异步处理主屏宿主指令，避免占用 Web STA。
    /// </summary>
    /// <param name="command">指令。</param>
    /// <returns>完成任务。</returns>
    private async Task HandleWebHostCommandAsync(string command)
    {
        try
        {
            if (command == "next")
            {
                await NextAsync().ConfigureAwait(false);
                return;
            }

            if (command == "prev")
            {
                await PreviousAsync().ConfigureAwait(false);
                return;
            }

            if (command == "stop")
            {
                await StopAsync().ConfigureAwait(false);
            }
        }
        catch
        {
            // 页面指令失败不打断放映
        }
    }

    /// <summary>
    /// 对单路会话执行前进一步或后退一步，失败不抛出。
    /// </summary>
    /// <param name="session">目标会话。</param>
    /// <param name="forward">true 为下一步。</param>
    private static void TryStep(IScreenPlayer session, bool forward)
    {
        try
        {
            if (forward)
            {
                session.Next();
            }
            else
            {
                session.Previous();
            }
        }
        catch
        {
            // 单路失败不阻断其余屏幕
        }
    }

    /// <summary>
    /// 生成「超前等待落后」的短提示。
    /// </summary>
    /// <param name="readable">已读到进度的各路。</param>
    /// <param name="pivot">本拍对齐的基准进度。</param>
    /// <param name="waitingAhead">true 表示超前屏在等待。</param>
    /// <returns>提示文案。</returns>
    private static string? BuildWaitMessage(
        IReadOnlyList<(IScreenPlayer Session, PlaybackProgress Progress)> readable,
        PlaybackProgress pivot,
        bool waitingAhead)
    {
        var waiting = readable
            .Where(item => waitingAhead
                ? item.Progress.CompareTo(pivot) > 0
                : item.Progress.CompareTo(pivot) < 0)
            .Select(item => item.Session.Screen.Label)
            .ToList();
        var moving = readable
            .Where(item => item.Progress.CompareTo(pivot) == 0)
            .Select(item => item.Session.Screen.Label)
            .ToList();

        if (waiting.Count == 0)
        {
            return null;
        }

        return waitingAhead
            ? $"{string.Join("、", waiting)} 等待对齐（{pivot}），仅 {string.Join("、", moving)} 前进一步"
            : $"{string.Join("、", waiting)} 等待对齐，仅 {string.Join("、", moving)} 后退一步";
    }

    /// <summary>
    /// 关闭 Web / WPS 会话。
    /// </summary>
    /// <returns>完成任务。</returns>
    private async Task TearDownAsync()
    {
        StopLayoutWatch();
        foreach (var session in _players.OfType<WebPlaybackSession>().ToList())
        {
            try
            {
                session.Stop();
            }
            catch
            {
                // 继续清理
            }
        }

        await _pump.InvokeAsync(() =>
        {
            foreach (var session in _players.OfType<WpsPlaybackSession>().ToList())
            {
                try
                {
                    session.Stop();
                }
                catch
                {
                    // 继续退出进程
                }
            }

            if (_wpsHost is not null)
            {
                try
                {
                    _wpsHost.Quit();
                }
                catch
                {
                    // 继续清理
                }

                _wpsHost = null;
            }
        }).ConfigureAwait(false);

        _players.Clear();
    }
}
