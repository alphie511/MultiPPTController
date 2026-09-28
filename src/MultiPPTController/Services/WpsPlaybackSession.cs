using MultiPPTController.Models;

namespace MultiPPTController.Services;

/// <summary>
/// 副屏 WPS 单路放映：一份文稿、一块屏，不再拆 MDI、不嵌主屏宿主。
/// </summary>
public sealed class WpsPlaybackSession : IScreenPlayer
{
    private const int PpShowTypeSpeaker = 1;
    private const int PpShowTypeWindow = 2;
    private const int MsoTrue = -1;
    private const int MsoFalse = 0;

    private dynamic? _presentation;
    private dynamic? _showWindow;
    private IntPtr _showHwnd;
    private IntPtr _hostHwnd;
    private bool _supportsClickIndex;

    /// <summary>
    /// 绑定片库条目与目标屏幕。
    /// </summary>
    /// <param name="deck">片库条目。</param>
    /// <param name="screen">输出屏幕。</param>
    /// <param name="openPath">实际打开的文件路径；同一 PPT 多屏时可为临时副本。</param>
    /// <param name="ownsTempCopy">结束时是否删除临时副本。</param>
    public WpsPlaybackSession(DeckItem deck, DisplayScreen screen, string? openPath = null, bool ownsTempCopy = false)
    {
        Deck = deck;
        Screen = screen;
        OpenPath = openPath ?? deck.FilePath;
        OwnsTempCopy = ownsTempCopy;
    }

    /// <summary>对应片库条目。</summary>
    public DeckItem Deck { get; }

    /// <summary>输出屏幕。</summary>
    public DisplayScreen Screen { get; }

    /// <summary>WPS 打开用的文件路径。</summary>
    public string OpenPath { get; }

    /// <summary>是否在关闭时删除临时副本。</summary>
    public bool OwnsTempCopy { get; }

    /// <summary>已捕获的放映窗口句柄；未捕获则为 Zero。</summary>
    public IntPtr ShowHwnd => _showHwnd;

    /// <summary>所属独立 WPS 进程 ID；0 表示不限定。</summary>
    public int EngineProcessId { get; set; }

    /// <summary>
    /// 诊断：本路 COM 放映窗与已绑句柄。
    /// </summary>
    /// <returns>单行状态。</returns>
    public string DescribeShow()
    {
        var comHwnd = IntPtr.Zero;
        var showCount = -1;
        try
        {
            if (_showWindow is not null)
            {
                comHwnd = new IntPtr(Convert.ToInt64(_showWindow.HWND));
            }
        }
        catch
        {
            // 无 HWND
        }

        try
        {
            showCount = Convert.ToInt32(_presentation!.Application.SlideShowWindows.Count);
        }
        catch
        {
            // 无集合
        }

        return $"{Deck.DisplayName} @{Screen.Label} show=0x{_showHwnd.ToInt64():X} '{NativeWindowHelper.GetTitle(_showHwnd)}' com=0x{comHwnd.ToInt64():X} '{NativeWindowHelper.GetTitle(comHwnd)}' slideShowWindows={showCount}";
    }

    /// <summary>
    /// 只读打开演示文稿并读取页数。
    /// </summary>
    /// <param name="application">KWPP.Application 实例。</param>
    public void Open(dynamic application)
    {
        _presentation = application.Presentations.Open(OpenPath, MsoTrue, MsoFalse, MsoTrue);
        try
        {
            Deck.SlideCount = (int)_presentation.Slides.Count;
        }
        catch
        {
            Deck.SlideCount = null;
        }
    }

    /// <summary>
    /// 以窗口放映启动，并把窗口铺到目标屏幕。
    /// </summary>
    /// <param name="occupiedByOthers">其它路已占用的窗口，避免两路绑到同一隐藏窗。</param>
    public void StartShow(IEnumerable<IntPtr>? occupiedByOthers = null)
    {
        if (_presentation is null)
        {
            throw new InvalidOperationException($"尚未打开：{Deck.DisplayName}");
        }

        var before = NativeWindowHelper.CaptureTopWindows();
        var settings = _presentation.SlideShowSettings;
        settings.ShowWithAnimation = MsoTrue;
        settings.ShowWithNarration = MsoTrue;
        try
        {
            settings.ShowScrollbar = MsoFalse;
        }
        catch
        {
            // 个别版本无此属性
        }

        try
        {
            settings.ShowType = PpShowTypeSpeaker;
            _showWindow = settings.Run();
        }
        catch
        {
            _showWindow = null;
        }

        if (_showWindow is null)
        {
            try
            {
                settings.ShowType = PpShowTypeWindow;
                _showWindow = settings.Run();
            }
            catch
            {
                _showWindow = null;
            }
        }

        BindShowFromCollection();
        Thread.Sleep(250);
        if (!TryBindSpeakerShow(before, occupiedByOthers))
        {
            TryForceSpeakerWithF5(occupiedByOthers);
        }

        if (!IsBoundToSpeakerShow())
        {
            ResolveShowHwnd(before, occupiedByOthers);
            RelayoutToScreen(occupiedByOthers, raiseZOrder: true);
        }

        NativeWindowHelper.FitToBounds(_showHwnd, Screen.Bounds, raiseZOrder: true, allowEditor: true);
        NativeWindowHelper.HideReadingChrome(_showHwnd);
        NativeWindowHelper.MinimizeEditorShells();
    }

    /// <summary>
    /// 优先绑真正的全屏放映窗并铺到副屏；没有像素则失败，改走文档框。
    /// </summary>
    /// <param name="before">Run 之前的窗口。</param>
    /// <param name="occupiedByOthers">其它路句柄。</param>
    /// <returns>已绑到有画面的放映窗则为 true。</returns>
    private bool TryBindSpeakerShow(
        IReadOnlyList<NativeWindowHelper.TopWindow> before,
        IEnumerable<IntPtr>? occupiedByOthers)
    {
        BindHwndFromCom();
        var hwnd = _showHwnd;
        if (!NativeWindowHelper.IsAlive(hwnd))
        {
            hwnd = NativeWindowHelper.FindSpeakerShowWindow(EngineProcessId, occupiedByOthers);
        }

        if (!NativeWindowHelper.IsAlive(hwnd))
        {
            foreach (var hidden in NativeWindowHelper.FindHiddenShowCandidates(EngineProcessId, occupiedByOthers))
            {
                NativeWindowHelper.RevealWindow(hidden);
                hwnd = hidden;
                break;
            }
        }

        if (!NativeWindowHelper.IsAlive(hwnd))
        {
            return false;
        }

        NativeWindowHelper.RevealWindow(hwnd);
        NativeWindowHelper.FitToBounds(hwnd, Screen.Bounds, raiseZOrder: true, allowEditor: true);
        var pixels = NativeWindowHelper.SamplePixels(hwnd);
        if (pixels.Colored < 8)
        {
            NativeWindowHelper.HideWindow(hwnd);
            return false;
        }

        _showHwnd = hwnd;
        NativeWindowHelper.HideReadingChrome(hwnd);
        return true;
    }

    /// <summary>
    /// 当前句柄是否为真正的 Slide Show 放映窗（不是阅读视图底栏那种）。
    /// </summary>
    /// <returns>标题像放映窗则为 true。</returns>
    private bool IsBoundToSpeakerShow()
    {
        if (!NativeWindowHelper.IsAlive(_showHwnd))
        {
            return false;
        }

        return NativeWindowHelper.IsSlideShowTitle(NativeWindowHelper.GetTitle(_showHwnd));
    }

    /// <summary>
    /// COM Run 落到阅读视图时，向编辑框补发 F5 再绑一次放映窗。
    /// </summary>
    /// <param name="occupiedByOthers">其它路句柄。</param>
    private void TryForceSpeakerWithF5(IEnumerable<IntPtr>? occupiedByOthers)
    {
        var editor = NativeWindowHelper.FindEditorFrame(OpenPath, occupiedByOthers);
        if (!NativeWindowHelper.IsAlive(editor))
        {
            editor = NativeWindowHelper.FindDocumentSurface(OpenPath, occupiedByOthers);
        }

        if (!NativeWindowHelper.IsAlive(editor))
        {
            return;
        }

        NativeWindowHelper.FocusReading(editor);
        NativeWindowHelper.PostVirtualKey(editor, 0x74);
        Thread.Sleep(400);
        TryBindSpeakerShow(NativeWindowHelper.CaptureTopWindows(), occupiedByOthers);
    }

    /// <summary>
    /// 读取当前页码与单击动画进度。
    /// </summary>
    /// <returns>进度快照；读失败时页码为 0。</returns>
    public PlaybackProgress ReadProgress()
    {
        RefreshShowWindow();
        if (_showWindow is null)
        {
            BindShowFromCollection();
        }

        if (_showWindow is null)
        {
            return default;
        }

        var slide = 0;
        var click = 0;
        try
        {
            slide = Convert.ToInt32(_showWindow!.View.CurrentShowPosition);
        }
        catch
        {
            try
            {
                slide = Convert.ToInt32(_showWindow!.View.Slide.SlideIndex);
            }
            catch
            {
                // 窗口尚未就绪或 COM 不支持
            }
        }

        try
        {
            click = Convert.ToInt32(_showWindow!.View.GetClickIndex());
            _supportsClickIndex = true;
        }
        catch
        {
            try
            {
                click = Convert.ToInt32(_showWindow!.View.ClickIndex);
                _supportsClickIndex = true;
            }
            catch
            {
                // WPS 个别版本没有单击进度接口，只按页同步
            }
        }

        return new PlaybackProgress(slide, click);
    }

    /// <summary>
    /// 开播后预热该路放映，避免用户第一次下一步被当成「激活窗口」吃掉。
    /// </summary>
    public void WarmUp()
    {
        NativeWindowHelper.FocusReading(_showHwnd);
    }

    /// <summary>
    /// 同步前进一步：先激活并等待窗口就绪，进度没变则再补发一次。
    /// </summary>
    public void Next()
    {
        StepAndConfirm(forward: true);
    }

    /// <summary>
    /// 同步后退一步。
    /// </summary>
    public void Previous()
    {
        StepAndConfirm(forward: false);
    }

    /// <summary>
    /// 激活 → 短等待 → 发指令；若能读到进度且没有变化，再补发一次（主屏第一下常被激活吞掉）。
    /// </summary>
    /// <param name="forward">true 为下一步。</param>
    private void StepAndConfirm(bool forward)
    {
        BindShowFromCollection();
        var before = ReadProgress();
        var comOk = TryStepShowView(forward);
        var after = ReadProgress();
        if (comOk && after.HasSlide && before.HasSlide && after.CompareTo(before) != 0)
        {
            return;
        }

        NativeWindowHelper.FocusReading(_showHwnd);
        SendWindowStep(forward);
        NativeWindowHelper.PostStepKeys(_showHwnd, _hostHwnd, forward ? 0x22 : 0x21);
    }

    /// <summary>
    /// 用 COM 放映视图前进一步或后退；HWND=0 时集合里仍可能有可用 View。
    /// </summary>
    /// <param name="forward">true 为下一步。</param>
    /// <returns>COM 调用成功则为 true。</returns>
    private bool TryStepShowView(bool forward)
    {
        if (TryInvokeView(view => StepView(view, forward)))
        {
            return true;
        }

        try
        {
            foreach (dynamic window in _presentation!.Application.SlideShowWindows)
            {
                StepView(window.View, forward);
                _showWindow = window;
                return true;
            }
        }
        catch
        {
            // 无集合或版本不支持
        }

        try
        {
            var live = _presentation!.SlideShowWindow;
            if (live is not null)
            {
                StepView(live.View, forward);
                _showWindow = live;
                return true;
            }
        }
        catch
        {
            // 无单窗属性
        }

        return false;
    }

    /// <summary>
    /// 对放映视图执行下一步或上一步。
    /// </summary>
    /// <param name="view">SlideShowView。</param>
    /// <param name="forward">true 为下一步。</param>
    private static void StepView(dynamic view, bool forward)
    {
        if (forward)
        {
            view.Next();
        }
        else
        {
            view.Previous();
        }
    }

    /// <summary>
    /// 只向本路宿主/画布发键，不用共享 COM（否则只会打动最后打开的副屏）。
    /// </summary>
    /// <param name="forward">true 为下一步。</param>
    private void SendWindowStep(bool forward)
    {
        NativeWindowHelper.PostStepKeys(_showHwnd, _hostHwnd, forward ? 0x27 : 0x25);
    }

    /// <summary>
    /// COM 放映对象是否带有本路独立 HWND。共享集合不能当进度源。
    /// </summary>
    /// <returns>有独立句柄则为 true。</returns>
    private bool HasDedicatedComShow()
    {
        if (_showWindow is null)
        {
            return false;
        }

        try
        {
            return new IntPtr(Convert.ToInt64(_showWindow.HWND)) != IntPtr.Zero;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// 结束该份放映并关闭文稿（不退出 WPS 进程）。
    /// </summary>
    public void Close()
    {
        NativeWindowHelper.DestroyPlaybackHost(_showHwnd, _hostHwnd);
        _hostHwnd = IntPtr.Zero;

        try
        {
            TryInvokeView(view => view.Exit());
        }
        catch
        {
            // 窗口可能已被用户关掉
        }

        try
        {
            _presentation?.Close();
        }
        catch
        {
            // 忽略重复关闭
        }

        _showWindow = null;
        _presentation = null;
        _showHwnd = IntPtr.Zero;
        Deck.Status = "已停止";

        if (OwnsTempCopy)
        {
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

    /// <inheritdoc />
    public void Stop() => Close();

    /// <summary>
    /// 按目标屏的真实像素重新铺满放映窗（避免 WPS 用主屏分辨率去套副屏）。
    /// </summary>
    /// <param name="occupiedByOthers">其它路已占用的窗口，避免绑错。</param>
    /// <param name="raiseZOrder">开播首次铺窗时置顶盖任务栏；翻页时不要置顶，以免盖住悬浮球。</param>
    public void RelayoutToScreen(IEnumerable<IntPtr>? occupiedByOthers = null, bool raiseZOrder = false)
    {
        if (IsBoundToSpeakerShow())
        {
            if (!NativeWindowHelper.MatchesBounds(_showHwnd, Screen.Bounds, 24))
            {
                NativeWindowHelper.FitToBounds(_showHwnd, Screen.Bounds, raiseZOrder);
            }

            NativeWindowHelper.HideReadingChrome(_showHwnd);
            return;
        }

        if (NativeWindowHelper.IsAlive(_showHwnd) &&
            NativeWindowHelper.MatchesBounds(_showHwnd, Screen.Bounds, 24))
        {
            NativeWindowHelper.HideReadingChrome(_showHwnd);
            return;
        }

        RefreshShowWindow();
        BindHwndFromCom();
        if (!IsHwndOwnedByEngine(_showHwnd) ||
            !NativeWindowHelper.IsUsableShowWindow(_showHwnd) ||
            IsOccupiedByOthers(_showHwnd, occupiedByOthers))
        {
            _showHwnd = IntPtr.Zero;
        }

        if (!NativeWindowHelper.IsAlive(_showHwnd))
        {
            _showHwnd = NativeWindowHelper.FindEditorFrame(OpenPath, occupiedByOthers);
        }

        if (!NativeWindowHelper.IsAlive(_showHwnd))
        {
            _showHwnd = NativeWindowHelper.FindDocumentSurface(OpenPath, occupiedByOthers);
        }

        if (!NativeWindowHelper.IsAlive(_showHwnd) && EngineProcessId > 0)
        {
            _showHwnd = NativeWindowHelper.FindWpsWindowByProcessId(EngineProcessId, occupiedByOthers);
        }

        ApplyLayout(_showHwnd, raiseZOrder);
    }

    /// <summary>
    /// 按本路进程锁定放映窗：WPS 总会先开在主屏，不能按「已经在副屏上」去找。
    /// </summary>
    /// <param name="before">Run 之前的窗口快照。</param>
    private void ResolveShowHwnd(
        IReadOnlyList<NativeWindowHelper.TopWindow> before,
        IEnumerable<IntPtr>? occupiedByOthers = null)
    {
        BindHwndFromCom();
        if (!IsHwndOwnedByEngine(_showHwnd) ||
            !NativeWindowHelper.IsUsableShowWindow(_showHwnd) ||
            IsOccupiedByOthers(_showHwnd, occupiedByOthers))
        {
            _showHwnd = IntPtr.Zero;
        }

        if (!NativeWindowHelper.IsAlive(_showHwnd))
        {
            _showHwnd = NativeWindowHelper.FindEditorFrame(OpenPath, occupiedByOthers);
        }

        if (!NativeWindowHelper.IsAlive(_showHwnd))
        {
            _showHwnd = NativeWindowHelper.FindDocumentSurface(OpenPath, occupiedByOthers);
        }

        if (!NativeWindowHelper.IsAlive(_showHwnd))
        {
            var after = NativeWindowHelper.CaptureTopWindows();
            var created = NativeWindowHelper.FindNewWpsWindow(before, after, EngineProcessId);
            if (created is { } window &&
                !IsOccupiedByOthers(window.Hwnd, occupiedByOthers) &&
                NativeWindowHelper.IsEditorFrameClass(NativeWindowHelper.GetClassName(window.Hwnd)))
            {
                _showHwnd = window.Hwnd;
            }
        }
    }

    /// <summary>
    /// Run() 在 WPS 365 常返回空，改从 SlideShowWindows 集合补绑。
    /// </summary>
    private void BindShowFromCollection()
    {
        if (_showWindow is not null || _presentation is null)
        {
            return;
        }

        try
        {
            var live = _presentation.SlideShowWindow;
            if (live is not null && new IntPtr(Convert.ToInt64(live.HWND)) != IntPtr.Zero)
            {
                _showWindow = live;
                return;
            }
        }
        catch
        {
            // 无单窗属性
        }

        try
        {
            foreach (dynamic window in _presentation.Application.SlideShowWindows)
            {
                _showWindow = window;
                return;
            }
        }
        catch
        {
            // 无集合
        }
    }

    /// <summary>
    /// 副屏只铺 PP12FrameClass 文档框，不再 SetParent 到 MultiPPT-host。
    /// </summary>
    /// <param name="hwnd">目标窗。</param>
    /// <param name="raiseZOrder">是否置顶。</param>
    private void ApplyLayout(IntPtr hwnd, bool raiseZOrder)
    {
        var frame = NativeWindowHelper.FindEditorFrame(OpenPath);
        if (NativeWindowHelper.IsAlive(frame))
        {
            NativeWindowHelper.FitEditorToBounds(frame, Screen.Bounds, raiseZOrder);
            NativeWindowHelper.HideReadingChrome(frame);
            _showHwnd = frame;
            return;
        }

        NativeWindowHelper.FitToBounds(hwnd, Screen.Bounds, raiseZOrder, allowEditor: true);
        NativeWindowHelper.HideReadingChrome(hwnd);
    }

    /// <summary>
    /// 句柄是否已被其它路占用。
    /// </summary>
    /// <param name="hwnd">窗口。</param>
    /// <param name="occupiedByOthers">其它路句柄。</param>
    /// <returns>已被占用则为 true。</returns>
    private static bool IsOccupiedByOthers(IntPtr hwnd, IEnumerable<IntPtr>? occupiedByOthers)
    {
        return hwnd != IntPtr.Zero &&
               occupiedByOthers is not null &&
               occupiedByOthers.Any(item => item == hwnd);
    }

    /// <summary>
    /// 句柄是否属于本路 WPS 进程。
    /// </summary>
    /// <param name="hwnd">窗口。</param>
    /// <returns>属于或未限定进程则为 true。</returns>
    private bool IsHwndOwnedByEngine(IntPtr hwnd)
    {
        if (!NativeWindowHelper.IsAlive(hwnd))
        {
            return false;
        }

        if (NativeWindowHelper.IsDocumentSurfaceClass(NativeWindowHelper.GetClassName(hwnd)))
        {
            return true;
        }

        return EngineProcessId <= 0 ||
               NativeWindowHelper.GetWindowProcessId(hwnd) == EngineProcessId;
    }

    /// <summary>
    /// 用该屏物理像素写入 COM 坐标，把放映从主屏挪到目标屏。
    /// </summary>
    private void PlaceByCom()
    {
        if (_showWindow is null)
        {
            return;
        }

        var bounds = NativeWindowHelper.GetPhysicalMonitorBounds(Screen.Bounds);
        var dpi = NativeWindowHelper.GetWindowDpi(_showHwnd);
        try
        {
            _showWindow.Left = bounds.Left * 72.0 / dpi;
            _showWindow.Top = bounds.Top * 72.0 / dpi;
            _showWindow.Width = bounds.Width * 72.0 / dpi;
            _showWindow.Height = bounds.Height * 72.0 / dpi;
        }
        catch
        {
            // 部分版本只读，交给 Win32
        }
    }

    /// <summary>
    /// 从 WPS 放映对象读取窗口句柄。
    /// </summary>
    private void BindHwndFromCom()
    {
        if (_showWindow is null)
        {
            return;
        }

        try
        {
            var hwnd = new IntPtr(Convert.ToInt64(_showWindow.HWND));
            if (IsHwndOwnedByEngine(hwnd) && NativeWindowHelper.IsUsableShowWindow(hwnd))
            {
                _showHwnd = hwnd;
            }
        }
        catch
        {
            // 个别版本没有 HWND 属性
        }
    }

    /// <summary>
    /// 重新取得「这一路」放映窗，禁止误绑到另一块屏正在活动的窗口。
    /// </summary>
    public void RefreshShowWindow()
    {
        if (_presentation is null)
        {
            return;
        }

        if (TryFindShowWindow() is { } matched)
        {
            _showWindow = matched;
            return;
        }

        if (_showWindow is not null)
        {
            return;
        }

        try
        {
            var live = _presentation.SlideShowWindow;
            if (live is not null)
            {
                _showWindow = live;
            }
        }
        catch
        {
            // 保持原引用
        }
    }

    /// <summary>
    /// 按本路句柄或所在屏幕找到对应的 SlideShowWindow。
    /// </summary>
    /// <returns>匹配的放映窗；找不到则为 null。</returns>
    private dynamic? TryFindShowWindow()
    {
        try
        {
            foreach (dynamic window in _presentation.Application.SlideShowWindows)
            {
                if (IsSessionShowWindow(window))
                {
                    return window;
                }
            }
        }
        catch
        {
            // 无集合
        }

        try
        {
            var live = _presentation.SlideShowWindow;
            if (live is not null && IsSessionShowWindow(live))
            {
                return live;
            }
        }
        catch
        {
            // 忽略
        }

        return null;
    }

    /// <summary>
    /// 该 COM 放映窗是否属于本路屏幕。
    /// </summary>
    /// <param name="window">SlideShowWindow。</param>
    /// <returns>属于本路则为 true。</returns>
    private bool IsSessionShowWindow(dynamic window)
    {
        try
        {
            var hwnd = new IntPtr(Convert.ToInt64(window.HWND));
            if (NativeWindowHelper.IsAlive(_showHwnd) && hwnd == _showHwnd)
            {
                return true;
            }

            return NativeWindowHelper.IsOnScreen(hwnd, Screen.Bounds);
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// 仅 COM 激活本路放映，不抢前台、不重铺，避免悬浮球闪烁。
    /// </summary>
    private void ActivateForCommand()
    {
        RefreshShowWindow();
        try
        {
            _showWindow?.Activate();
        }
        catch
        {
            // 部分版本无 Activate
        }
    }

    /// <summary>
    /// 开播预热：激活并校正布局。
    /// </summary>
    public void Activate()
    {
        ActivateForCommand();
        RelayoutToScreen();
    }

    /// <summary>
    /// 对放映视图执行操作。
    /// </summary>
    /// <param name="action">视图回调。</param>
    /// <returns>COM 调用成功则为 true。</returns>
    private bool TryInvokeView(Action<dynamic> action)
    {
        if (_showWindow is null)
        {
            return false;
        }

        try
        {
            action(_showWindow.View);
            return true;
        }
        catch
        {
            return false;
        }
    }

}
