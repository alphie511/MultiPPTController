using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

namespace MultiPPTController.Services;

/// <summary>
/// 捕获 WPS 放映窗口并用 Win32 铺满指定屏幕。
/// </summary>
public static class NativeWindowHelper
{
    private const uint SwpShowWindow = 0x0040;
    private const uint SwpFrameChanged = 0x0020;
    private const uint SwpNoCopyBits = 0x0100;
    private const uint SwpNoZOrder = 0x0004;
    private const uint SwpNoSendChanging = 0x0400;
    private const uint MonitorDefaultToNearest = 2;
    private const uint ObjIdNativeOm = 0xFFFFFFF0;
    private static readonly IntPtr HwndTopMost = new(-1);
    private const int SwHide = 0;
    private const int SwShow = 5;
    private const int SwMinimize = 6;
    private const int SwRestore = 9;
    private const int SwShowNoActivate = 8;
    private const uint SwpNoMove = 0x0002;
    private const uint SwpNoSize = 0x0001;
    private const uint SwpNoActivate = 0x0010;
    private const int GwlExStyle = -20;
    private const int WsExNoActivate = 0x08000000;
    private const int WsExToolWindow = 0x00000080;
    private const int WmKeyDown = 0x0100;
    private const int WmKeyUp = 0x0101;
    private const int WmSize = 0x0005;
    private const int WmSetFocus = 0x0007;
    private const int WmMouseActivate = 0x0021;
    private const uint RdwInvalidate = 0x0001;
    private const uint RdwAllChildren = 0x0080;
    private const uint RdwUpdateNow = 0x0100;
    private const int GwlStyle = -16;
    private const int WsCaption = 0x00C00000;
    private const int WsThickFrame = 0x00040000;
    private const int WsBorder = 0x00800000;
    private const int WsDlgFrame = 0x00400000;
    private const int WsChild = 0x40000000;
    private const int WsPopup = unchecked((int)0x80000000);
    private const int WsVisible = 0x10000000;

    /// <summary>
    /// 可见顶层窗口快照。
    /// </summary>
    /// <param name="Hwnd">窗口句柄。</param>
    /// <param name="ProcessName">进程名（小写，无扩展名）。</param>
    /// <param name="Title">窗口标题。</param>
    /// <param name="Width">宽度像素。</param>
    /// <param name="Height">高度像素。</param>
    public readonly record struct TopWindow(
        IntPtr Hwnd,
        string ProcessName,
        string Title,
        int Left,
        int Top,
        int Width,
        int Height);

    /// <summary>
    /// 含隐藏态的 WPS 顶层窗快照。
    /// </summary>
    /// <param name="Hwnd">窗口句柄。</param>
    /// <param name="ProcessId">进程 ID。</param>
    /// <param name="ProcessName">进程名。</param>
    /// <param name="Title">窗口标题。</param>
    /// <param name="ClassName">窗口类名。</param>
    /// <param name="Left">左。</param>
    /// <param name="Top">上。</param>
    /// <param name="Width">宽。</param>
    /// <param name="Height">高。</param>
    /// <param name="Visible">是否可见。</param>
    public readonly record struct ProcessWindow(
        IntPtr Hwnd,
        int ProcessId,
        string ProcessName,
        string Title,
        string ClassName,
        int Left,
        int Top,
        int Width,
        int Height,
        bool Visible);

    /// <summary>
    /// 枚举当前可见且面积足够的顶层窗口。
    /// </summary>
    /// <returns>窗口快照列表。</returns>
    public static IReadOnlyList<TopWindow> CaptureTopWindows()
    {
        var result = new List<TopWindow>();
        EnumWindows((hwnd, _) =>
        {
            if (!IsWindowVisible(hwnd))
            {
                return true;
            }

            GetWindowRect(hwnd, out var rect);
            var width = rect.Right - rect.Left;
            var height = rect.Bottom - rect.Top;
            if (width < 160 || height < 120)
            {
                return true;
            }

            GetWindowThreadProcessId(hwnd, out var pid);
            string processName;
            try
            {
                processName = Process.GetProcessById((int)pid).ProcessName.ToLowerInvariant();
            }
            catch
            {
                return true;
            }

            var title = new StringBuilder(512);
            GetWindowText(hwnd, title, title.Capacity);
            result.Add(new TopWindow(hwnd, processName, title.ToString(), rect.Left, rect.Top, width, height));
            return true;
        }, IntPtr.Zero);

        return result;
    }

    /// <summary>
    /// 枚举全部 WPS/wpp 顶层窗（含隐藏、小窗），供诊断。
    /// </summary>
    /// <returns>窗口快照。</returns>
    public static IReadOnlyList<ProcessWindow> CaptureAllWpsWindows()
    {
        var result = new List<ProcessWindow>();
        EnumWindows((hwnd, _) =>
        {
            GetWindowThreadProcessId(hwnd, out var pid);
            string processName;
            try
            {
                processName = Process.GetProcessById((int)pid).ProcessName.ToLowerInvariant();
            }
            catch
            {
                return true;
            }

            if (!IsWpsProcess(processName) &&
                !IsPlaybackHostTitle(GetTitle(hwnd)) &&
                !IsWebPlayerTitle(GetTitle(hwnd)))
            {
                return true;
            }

            GetWindowRect(hwnd, out var rect);
            var title = new StringBuilder(512);
            GetWindowText(hwnd, title, title.Capacity);
            result.Add(new ProcessWindow(
                hwnd,
                (int)pid,
                processName,
                title.ToString(),
                GetClassName(hwnd),
                rect.Left,
                rect.Top,
                Math.Max(0, rect.Right - rect.Left),
                Math.Max(0, rect.Bottom - rect.Top),
                IsWindowVisible(hwnd)));
            return true;
        }, IntPtr.Zero);
        return result;
    }

    /// <summary>
    /// 从两次快照中找出新增的 WPS 相关窗口，取面积最大者。
    /// </summary>
    /// <param name="before">Run 之前的窗口。</param>
    /// <param name="after">Run 之后的窗口。</param>
    /// <returns>最可能的放映窗口；找不到则为 null。</returns>
    public static TopWindow? FindNewWpsWindow(
        IReadOnlyList<TopWindow> before,
        IReadOnlyList<TopWindow> after,
        int processId = 0)
    {
        var known = before.Select(item => item.Hwnd).ToHashSet();
        return after
            .Where(item => !known.Contains(item.Hwnd) && IsWpsProcess(item.ProcessName))
            .Where(item => processId <= 0 || GetWindowProcessId(item.Hwnd) == processId)
            .Where(item => !IsEditorTitle(item.Title))
            .OrderByDescending(item => IsSlideShowTitle(item.Title) ? 1 : 0)
            .ThenByDescending(item => item.Width * (long)item.Height)
            .Select(item => (TopWindow?)item)
            .FirstOrDefault();
    }

    /// <summary>
    /// 读取窗口所属进程 ID。
    /// </summary>
    /// <param name="hwnd">窗口句柄。</param>
    /// <returns>进程 ID；失败则为 0。</returns>
    public static int GetWindowProcessId(IntPtr hwnd)
    {
        if (hwnd == IntPtr.Zero)
        {
            return 0;
        }

        GetWindowThreadProcessId(hwnd, out var pid);
        return (int)pid;
    }

    /// <summary>
    /// 是否为 WPS 365 文档框（幻灯片画在这个框的 Reading 子窗里）。
    /// </summary>
    /// <param name="className">窗口类名。</param>
    /// <returns>是文档框则为 true。</returns>
    public static bool IsEditorFrameClass(string className)
    {
        return className.Equals("PP12FrameClass", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// 是否为程序自建的放映宿主标题。
    /// </summary>
    /// <param name="title">标题。</param>
    /// <returns>是宿主则为 true。</returns>
    public static bool IsPlaybackHostTitle(string title)
    {
        return title.Equals("MultiPPT-host", StringComparison.Ordinal);
    }

    /// <summary>
    /// 是否为主屏 WebView2 放映窗标题。
    /// </summary>
    /// <param name="title">标题。</param>
    /// <returns>是 Web 放映窗则为 true。</returns>
    public static bool IsWebPlayerTitle(string title)
    {
        return title.Equals("MultiPPT-web", StringComparison.Ordinal);
    }

    /// <summary>
    /// 按文稿文件名查找可见的 WPS 文档框。
    /// </summary>
    /// <param name="filePath">打开路径或文件名。</param>
    /// <param name="exclude">其它路已占用的窗口。</param>
    /// <returns>文档框句柄；找不到则为 Zero。</returns>
    public static IntPtr FindEditorFrame(string filePath, IEnumerable<IntPtr>? exclude = null)
    {
        var taken = exclude?.Where(item => item != IntPtr.Zero).ToHashSet() ?? [];
        var names = new[]
        {
            Path.GetFileName(filePath),
            Path.GetFileNameWithoutExtension(filePath)
        }.Where(item => !string.IsNullOrWhiteSpace(item)).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();

        IntPtr fallback = IntPtr.Zero;
        var fallbackArea = 0L;
        foreach (var window in CaptureAllWpsWindows())
        {
            if (taken.Contains(window.Hwnd) || !IsEditorFrameClass(window.ClassName))
            {
                continue;
            }

            var area = window.Width * (long)window.Height;
            if (names.Any(name => window.Title.Contains(name, StringComparison.OrdinalIgnoreCase)))
            {
                return window.Hwnd;
            }

            if (!window.Visible)
            {
                continue;
            }

            if (area > fallbackArea)
            {
                fallbackArea = area;
                fallback = window.Hwnd;
            }
        }

        return fallback;
    }

    /// <summary>
    /// 按文稿名查找 MDI 文档页（WPS 365 多份 PPT 叠在同一外壳里）。
    /// </summary>
    /// <param name="filePath">打开路径。</param>
    /// <param name="exclude">其它路已占用的窗口。</param>
    /// <returns>mdiClass 或文档框；找不到则为 Zero。</returns>
    public static IntPtr FindDocumentSurface(string filePath, IEnumerable<IntPtr>? exclude = null)
    {
        var taken = exclude?.Where(item => item != IntPtr.Zero).ToHashSet() ?? [];
        var fileName = Path.GetFileName(filePath);
        var stem = Path.GetFileNameWithoutExtension(filePath);
        var best = IntPtr.Zero;
        var bestScore = int.MinValue;

        foreach (var root in CaptureAllWpsWindows())
        {
            foreach (var hwnd in CollectWindowTree(root.Hwnd))
            {
                if (taken.Contains(hwnd) || !IsAlive(hwnd))
                {
                    continue;
                }

                var className = GetClassName(hwnd);
                var title = GetTitle(hwnd);
                if (!IsDocumentSurfaceClass(className))
                {
                    continue;
                }

                if (!title.Contains(stem, StringComparison.OrdinalIgnoreCase) &&
                    (string.IsNullOrWhiteSpace(fileName) || !title.Contains(fileName, StringComparison.OrdinalIgnoreCase)))
                {
                    continue;
                }

                var score = 0;
                if (className.Equals("mdiClass", StringComparison.OrdinalIgnoreCase))
                {
                    score += 20;
                }

                if (!title.Contains(" : ", StringComparison.Ordinal))
                {
                    score += 8;
                }

                if (IsWindowVisible(hwnd))
                {
                    score += 3;
                }

                if (score > bestScore)
                {
                    bestScore = score;
                    best = hwnd;
                }
            }
        }

        return best != IntPtr.Zero ? best : FindEditorFrame(filePath, exclude);
    }

    /// <summary>
    /// 是否为可铺屏的文档表面。
    /// </summary>
    /// <param name="className">类名。</param>
    /// <returns>是文档框或 MDI 页则为 true。</returns>
    public static bool IsDocumentSurfaceClass(string className)
    {
        return IsEditorFrameClass(className) ||
               className.Equals("mdiClass", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// 创建铺满目标屏的宿主窗，用来托住从 WPS 拆出的文档页。
    /// </summary>
    /// <param name="bounds">目标屏幕。</param>
    /// <returns>宿主句柄；失败则为 Zero。</returns>
    public static IntPtr CreatePlaybackHost(System.Drawing.Rectangle bounds)
    {
        bounds = GetPhysicalMonitorBounds(bounds);
        var host = CreateWindowEx(
            0x00000008,
            "Static",
            "MultiPPT-host",
            WsPopup | WsVisible,
            bounds.X,
            bounds.Y,
            bounds.Width,
            bounds.Height,
            IntPtr.Zero,
            IntPtr.Zero,
            GetModuleHandle(null),
            IntPtr.Zero);
        if (host != IntPtr.Zero)
        {
            SetWindowPos(host, HwndTopMost, bounds.X, bounds.Y, bounds.Width, bounds.Height, SwpShowWindow | SwpFrameChanged | SwpNoSendChanging);
        }

        return host;
    }

    /// <summary>
    /// 把文档页嵌进宿主并撑满。WPS 会钳制第二扇独立窗的尺寸，嵌进自有宿主才能铺满副屏。
    /// </summary>
    /// <param name="content">mdiClass / 文档框。</param>
    /// <param name="host">宿主窗。</param>
    /// <param name="bounds">目标屏幕。</param>
    public static void EmbedInHost(IntPtr content, IntPtr host, System.Drawing.Rectangle bounds)
    {
        if (!IsAlive(content) || !IsAlive(host))
        {
            return;
        }

        bounds = GetPhysicalMonitorBounds(bounds);
        SetWindowPos(host, HwndTopMost, bounds.X, bounds.Y, bounds.Width, bounds.Height, SwpShowWindow | SwpFrameChanged | SwpNoSendChanging);
        var style = GetWindowLong(content, GwlStyle);
        SetWindowLong(content, GwlStyle, (style & ~WsPopup & ~WsCaption & ~WsThickFrame) | WsChild | WsVisible);
        SetParent(content, host);
        ForceFullResolution(content, host, bounds);
        ShowWindow(host, SwShow);
        ShowWindow(content, SwShow);
        FocusReading(content);
    }

    /// <summary>
    /// 先微缩再拉满，逼 WPS 按目标屏像素重绘，避免第一页停在小画布的糊图。
    /// </summary>
    /// <param name="content">文档页。</param>
    /// <param name="host">宿主。</param>
    /// <param name="bounds">目标屏幕。</param>
    public static void ForceFullResolution(IntPtr content, IntPtr host, System.Drawing.Rectangle bounds)
    {
        if (!IsAlive(content) || !IsAlive(host))
        {
            return;
        }

        bounds = GetPhysicalMonitorBounds(bounds);
        SetWindowPos(host, HwndTopMost, bounds.X, bounds.Y, bounds.Width, bounds.Height, SwpShowWindow | SwpFrameChanged | SwpNoSendChanging);
        MoveWindow(content, 0, 0, bounds.Width, bounds.Height, true);
        SetWindowPos(content, IntPtr.Zero, 0, 0, bounds.Width, bounds.Height, SwpNoZOrder | SwpShowWindow | SwpFrameChanged | SwpNoSendChanging);
        FillReadingTree(content, bounds.Width, bounds.Height);
    }

    /// <summary>
    /// 把输入焦点落到 Reading 画布，不抢系统前台，避免第一步被当成「点选窗口」。
    /// </summary>
    /// <param name="content">文档页。</param>
    public static void FocusReading(IntPtr content)
    {
        var target = FindReadingSurface(content);
        if (!IsAlive(target))
        {
            return;
        }

        PostMessage(target, (uint)WmMouseActivate, target, IntPtr.Zero);
        PostMessage(target, (uint)WmSetFocus, IntPtr.Zero, IntPtr.Zero);
    }

    /// <summary>
    /// 向文档页、画布和宿主同时投递方向键，避免只打到副屏活动窗。
    /// </summary>
    /// <param name="content">文档页。</param>
    /// <param name="host">宿主；没有则为 Zero。</param>
    /// <param name="virtualKey">虚拟键。</param>
    public static void PostStepKeys(IntPtr content, IntPtr host, int virtualKey)
    {
        var reading = FindReadingSurface(content);
        foreach (var hwnd in new[] { reading, content, host }.Distinct())
        {
            if (IsAlive(hwnd))
            {
                PostVirtualKey(hwnd, virtualKey);
            }
        }
    }

    /// <summary>
    /// 结束放映时把文档页从宿主拆回，再销毁宿主。
    /// </summary>
    /// <param name="content">文档页。</param>
    /// <param name="host">宿主。</param>
    public static void DestroyPlaybackHost(IntPtr content, IntPtr host)
    {
        if (IsAlive(content) && IsAlive(host) && GetAncestorParent(content) == host)
        {
            SetParent(content, IntPtr.Zero);
        }

        if (IsAlive(host))
        {
            DestroyWindow(host);
        }
    }

    /// <summary>
    /// 把文档框铺到目标屏，并撑开内部 Reading / MDI 画布。
    /// </summary>
    /// <param name="hwnd">PP12FrameClass 文档框。</param>
    /// <param name="bounds">目标屏幕矩形。</param>
    /// <param name="raiseZOrder">是否置顶。</param>
    public static void FitEditorToBounds(IntPtr hwnd, System.Drawing.Rectangle bounds, bool raiseZOrder = true)
    {
        if (!IsAlive(hwnd) || bounds.Width <= 0 || bounds.Height <= 0)
        {
            return;
        }

        bounds = GetPhysicalMonitorBounds(bounds);
        if (MatchesBounds(hwnd, bounds, 24))
        {
            return;
        }

        var style = GetWindowLong(hwnd, GwlStyle);
        SetWindowLong(hwnd, GwlStyle, style & ~WsCaption & ~WsThickFrame & ~WsBorder & ~WsDlgFrame);

        var flags = SwpShowWindow | SwpFrameChanged | SwpNoCopyBits | SwpNoSendChanging | SwpNoActivate;
        var insertAfter = raiseZOrder ? HwndTopMost : IntPtr.Zero;
        if (!raiseZOrder)
        {
            flags |= SwpNoZOrder;
        }

        MoveWindow(hwnd, bounds.X, bounds.Y, bounds.Width, bounds.Height, false);
        SetWindowPos(hwnd, insertAfter, bounds.X, bounds.Y, bounds.Width, bounds.Height, flags);
        FillReadingTree(hwnd, bounds.Width, bounds.Height);
    }

    /// <summary>
    /// 在文档框里找到 Reading 画布，供按键落到真正的幻灯片。
    /// </summary>
    /// <param name="editor">文档框。</param>
    /// <returns>Reading 子窗或文档框本身。</returns>
    public static IntPtr FindReadingSurface(IntPtr editor)
    {
        if (!IsAlive(editor))
        {
            return IntPtr.Zero;
        }

        var found = IntPtr.Zero;
        EnumChildWindows(editor, (child, _) =>
        {
            var title = GetTitle(child);
            if (title.Contains("Reading", StringComparison.OrdinalIgnoreCase) ||
                title.Contains("阅读", StringComparison.OrdinalIgnoreCase))
            {
                found = child;
                return false;
            }

            return true;
        }, IntPtr.Zero);
        return found == IntPtr.Zero ? editor : found;
    }

    /// <summary>
    /// 隐藏空的 wpp Qt 罩层和阴影边框，避免挡住文档框。
    /// </summary>
    public static void HideEmptyEngineOverlays()
    {
        foreach (var window in CaptureAllWpsWindows())
        {
            if (!window.Visible)
            {
                continue;
            }

            if (window.ClassName.Contains("KLiteMainWindowShadowBorder", StringComparison.OrdinalIgnoreCase) ||
                (string.Equals(window.ProcessName, "wpp", StringComparison.OrdinalIgnoreCase) &&
                 window.ClassName.Contains("Qt5QWindow", StringComparison.OrdinalIgnoreCase) &&
                 string.IsNullOrWhiteSpace(window.Title)))
            {
                ShowWindow(window.Hwnd, SwHide);
            }
        }
    }

    /// <summary>
    /// 子文档拆走后，隐藏空的 WPS 外壳，避免挡住放映。
    /// </summary>
    public static void HideOrphanEditorShells()
    {
        foreach (var window in CaptureAllWpsWindows())
        {
            if (!window.Visible ||
                !IsEditorFrameClass(window.ClassName) ||
                !window.Title.Contains("WPS Office", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var hasDocument = false;
            EnumChildWindows(window.Hwnd, (child, _) =>
            {
                if (GetClassName(child).Equals("mdiClass", StringComparison.OrdinalIgnoreCase))
                {
                    hasDocument = true;
                    return false;
                }

                return true;
            }, IntPtr.Zero);

            if (!hasDocument)
            {
                ShowWindow(window.Hwnd, SwHide);
            }
        }
    }

    /// <summary>
    /// 撑开文档框内的幻灯片画布（不要无差别拉伸所有子窗）。
    /// </summary>
    /// <param name="parent">文档框或中间容器。</param>
    /// <param name="width">目标宽。</param>
    /// <param name="height">目标高。</param>
    private static void FillReadingTree(IntPtr parent, int width, int height)
    {
        EnumChildWindows(parent, (child, _) =>
        {
            var className = GetClassName(child);
            var title = GetTitle(child);
            if (IsEditorFrameClass(className) ||
                className.Equals("MDIClient", StringComparison.OrdinalIgnoreCase) ||
                className.Equals("mdiClass", StringComparison.OrdinalIgnoreCase) ||
                title.Contains("Reading", StringComparison.OrdinalIgnoreCase) ||
                title.Contains("阅读", StringComparison.OrdinalIgnoreCase))
            {
                SetWindowPos(child, IntPtr.Zero, 0, 0, width, height, SwpNoZOrder | SwpShowWindow | SwpFrameChanged);
                FillReadingTree(child, width, height);
            }

            return true;
        }, IntPtr.Zero);
    }

    /// <summary>
    /// 通知窗口尺寸已变，促使 WPS 按新像素重绘幻灯片。
    /// </summary>
    /// <param name="hwnd">窗口。</param>
    /// <param name="width">宽。</param>
    /// <param name="height">高。</param>
    private static void NotifySized(IntPtr hwnd, int width, int height)
    {
        if (!IsAlive(hwnd))
        {
            return;
        }

        var packed = (IntPtr)((height << 16) | (width & 0xFFFF));
        SendMessage(hwnd, WmSize, IntPtr.Zero, packed);
    }

    /// <summary>
    /// 立即重绘窗口树。
    /// </summary>
    /// <param name="hwnd">窗口。</param>
    private static void RedrawTree(IntPtr hwnd)
    {
        if (!IsAlive(hwnd))
        {
            return;
        }

        RedrawWindow(hwnd, IntPtr.Zero, IntPtr.Zero, RdwInvalidate | RdwAllChildren | RdwUpdateNow);
    }

    /// <summary>
    /// 采样窗口是否画出了非灰白内容（用于判断全屏是否空白）。
    /// </summary>
    /// <param name="hwnd">窗口。</param>
    /// <returns>有内容的采样点数 / 总采样；失败则为 0/0。</returns>
    public static (int Colored, int Samples) SamplePixels(IntPtr hwnd)
    {
        if (!IsAlive(hwnd))
        {
            return (0, 0);
        }

        GetWindowRect(hwnd, out var rect);
        var width = Math.Max(0, rect.Right - rect.Left);
        var height = Math.Max(0, rect.Bottom - rect.Top);
        if (width < 16 || height < 16 || width > 4000 || height > 3000)
        {
            return (0, 0);
        }

        try
        {
            using var bmp = new System.Drawing.Bitmap(width, height);
            using var graphics = System.Drawing.Graphics.FromImage(bmp);
            var hdc = graphics.GetHdc();
            try
            {
                PrintWindow(hwnd, hdc, 2);
            }
            finally
            {
                graphics.ReleaseHdc(hdc);
            }

            var colored = 0;
            var samples = 0;
            var stepX = Math.Max(8, width / 20);
            var stepY = Math.Max(8, height / 20);
            for (var y = 8; y < height; y += stepY)
            {
                for (var x = 8; x < width; x += stepX)
                {
                    var color = bmp.GetPixel(x, y);
                    samples++;
                    var sum = color.R + color.G + color.B;
                    var grey = Math.Abs(color.R - color.G) < 12 && Math.Abs(color.G - color.B) < 12;
                    if (sum > 40 && !(grey && color.R > 180))
                    {
                        colored++;
                    }
                }
            }

            return (colored, samples);
        }
        catch
        {
            return (0, 0);
        }
    }

    /// <summary>
    /// 找出被 WPS 藏起来的放映窗（本机 W365 的 SlideShow 在 wpp.exe 里且 Invisible）。
    /// </summary>
    /// <param name="processId">优先匹配的 wpp 进程；0 则不限。</param>
    /// <param name="exclude">其它路已占用的句柄。</param>
    /// <returns>按面积从大到小的候选句柄。</returns>
    public static IReadOnlyList<IntPtr> FindHiddenShowCandidates(int processId, IEnumerable<IntPtr>? exclude = null)
    {
        var taken = exclude?.Where(item => item != IntPtr.Zero).ToHashSet() ?? [];
        return CaptureAllWpsWindows()
            .Where(item => !taken.Contains(item.Hwnd))
            .Where(item => !item.Visible)
            .Where(item => !IsEditorTitle(item.Title))
            .Where(item => item.Width >= 400 && item.Height >= 280)
            .Where(item => processId <= 0 || item.ProcessId == processId || item.ProcessName == "wpp")
            .OrderByDescending(item => item.Width * (long)item.Height)
            .Select(item => item.Hwnd)
            .ToList();
    }

    /// <summary>
    /// 把隐藏/最小化的放映窗显示出来（不抢前台）。
    /// </summary>
    /// <param name="hwnd">窗口。</param>
    public static void RevealWindow(IntPtr hwnd)
    {
        if (!IsAlive(hwnd))
        {
            return;
        }

        ShowWindow(hwnd, SwRestore);
        ShowWindow(hwnd, SwShowNoActivate);
    }

    /// <summary>
    /// 隐藏空罩层或失败的放映窗，避免挡住副屏。
    /// </summary>
    /// <param name="hwnd">窗口。</param>
    public static void HideWindow(IntPtr hwnd)
    {
        if (IsAlive(hwnd))
        {
            ShowWindow(hwnd, SwHide);
        }
    }

    /// <summary>
    /// 查找真正的全屏放映窗（含 wpp 无标题 Qt 窗），排除编辑壳。
    /// </summary>
    /// <param name="processId">优先匹配的进程；0 则不限。</param>
    /// <param name="exclude">其它路已占用的句柄。</param>
    /// <returns>最可能的放映窗；找不到则为 Zero。</returns>
    public static IntPtr FindSpeakerShowWindow(int processId, IEnumerable<IntPtr>? exclude = null)
    {
        var taken = exclude?.Where(item => item != IntPtr.Zero).ToHashSet() ?? [];
        IntPtr best = IntPtr.Zero;
        var bestArea = 0L;
        foreach (var window in CaptureAllWpsWindows())
        {
            if (taken.Contains(window.Hwnd) ||
                IsEditorTitle(window.Title) ||
                IsEditorFrameClass(window.ClassName))
            {
                continue;
            }

            if (processId > 0 &&
                window.ProcessId != processId &&
                !string.Equals(window.ProcessName, "wpp", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var slideshow = IsSlideShowTitle(window.Title) ||
                            window.ClassName.Contains("Qt5QWindow", StringComparison.OrdinalIgnoreCase) ||
                            (string.IsNullOrWhiteSpace(window.Title) &&
                             window.Width >= 800 &&
                             window.Height >= 600);
            if (!slideshow)
            {
                continue;
            }

            var area = window.Width * (long)window.Height;
            if (area > bestArea)
            {
                bestArea = area;
                best = window.Hwnd;
            }
        }

        return best;
    }

    /// <summary>
    /// 隐藏阅读视图底栏 / 状态条，避免「幻灯片 1/23」挡住满屏。
    /// </summary>
    /// <param name="root">文档框或放映窗。</param>
    public static void HideReadingChrome(IntPtr root)
    {
        if (!IsAlive(root))
        {
            return;
        }

        EnumChildWindows(root, (child, _) =>
        {
            var title = GetTitle(child);
            var className = GetClassName(child);
            GetWindowRect(child, out var rect);
            var height = Math.Max(0, rect.Bottom - rect.Top);
            var width = Math.Max(0, rect.Right - rect.Left);
            var reading = title.Contains("Reading", StringComparison.OrdinalIgnoreCase) ||
                          title.Contains("阅读", StringComparison.OrdinalIgnoreCase) ||
                          className.Equals("mdiClass", StringComparison.OrdinalIgnoreCase);
            if (reading)
            {
                return true;
            }

            var status = title.Contains("幻灯片", StringComparison.Ordinal) ||
                         className.Contains("Status", StringComparison.OrdinalIgnoreCase) ||
                         className.Contains("ReBar", StringComparison.OrdinalIgnoreCase) ||
                         className.Contains("ToolBar", StringComparison.OrdinalIgnoreCase) ||
                         className.Contains("msctls_statusbar", StringComparison.OrdinalIgnoreCase);
            var thinBar = height is > 0 and <= 88 && width >= 400;
            if (status || thinBar)
            {
                ShowWindow(child, SwHide);
            }

            return true;
        }, IntPtr.Zero);
    }

    /// <summary>
    /// 读取窗口类名。
    /// </summary>
    /// <param name="hwnd">窗口。</param>
    /// <returns>类名；失败则为空串。</returns>
    public static string GetClassName(IntPtr hwnd)
    {
        if (!IsAlive(hwnd))
        {
            return string.Empty;
        }

        var buffer = new StringBuilder(256);
        QueryClassName(hwnd, buffer, buffer.Capacity);
        return buffer.ToString();
    }

    /// <summary>
    /// 按进程查找可见的放映标题窗，取面积最大者。
    /// </summary>
    /// <param name="processId">WPS 进程 ID。</param>
    /// <returns>窗口句柄；找不到则为 Zero。</returns>
    public static IntPtr FindWpsWindowByProcessId(int processId, IEnumerable<IntPtr>? exclude = null)
    {
        if (processId <= 0)
        {
            return IntPtr.Zero;
        }

        var taken = exclude?.Where(item => item != IntPtr.Zero).ToHashSet() ?? [];
        var best = IntPtr.Zero;
        var bestArea = 0L;
        foreach (var window in CaptureTopWindows())
        {
            if (taken.Contains(window.Hwnd) ||
                !IsWpsProcess(window.ProcessName) ||
                GetWindowProcessId(window.Hwnd) != processId ||
                IsEditorTitle(window.Title) ||
                !IsLikelyShowWindow(window.Title, GetClassName(window.Hwnd), window.ProcessName))
            {
                continue;
            }

            var area = window.Width * (long)window.Height;
            if (area > bestArea)
            {
                bestArea = area;
                best = window.Hwnd;
            }
        }

        return best;
    }

    /// <summary>
    /// 枚举该进程全部顶层窗口（含隐藏、小窗），供 NativeOM 绑定到新进程。
    /// </summary>
    /// <param name="processId">目标进程 ID。</param>
    /// <returns>顶层窗口句柄，可见大窗优先。</returns>
    public static IReadOnlyList<IntPtr> EnumerateProcessTopWindows(int processId)
    {
        var scored = new List<(IntPtr Hwnd, long Score)>();
        if (processId <= 0)
        {
            return [];
        }

        EnumWindows((hwnd, _) =>
        {
            GetWindowThreadProcessId(hwnd, out var pid);
            if ((int)pid != processId)
            {
                return true;
            }

            GetWindowRect(hwnd, out var rect);
            var width = Math.Max(0, rect.Right - rect.Left);
            var height = Math.Max(0, rect.Bottom - rect.Top);
            var visible = IsWindowVisible(hwnd);
            var score = width * (long)height + (visible ? 1_000_000_000L : 0);
            scored.Add((hwnd, score));
            return true;
        }, IntPtr.Zero);

        return scored
            .OrderByDescending(item => item.Score)
            .Select(item => item.Hwnd)
            .ToList();
    }

    /// <summary>
    /// 收集窗口及其全部子窗口。
    /// </summary>
    /// <param name="root">顶层窗口。</param>
    /// <returns>自身 + 子窗口句柄。</returns>
    public static IReadOnlyList<IntPtr> CollectWindowTree(IntPtr root)
    {
        var list = new List<IntPtr>();
        if (root == IntPtr.Zero)
        {
            return list;
        }

        list.Add(root);
        EnumChildWindows(root, (child, _) =>
        {
            list.Add(child);
            return true;
        }, IntPtr.Zero);
        return list;
    }

    /// <summary>
    /// 读取窗口 DPI，用于把屏幕像素换成 WPS 磅值。
    /// </summary>
    /// <param name="hwnd">窗口。</param>
    /// <returns>DPI；失败则为 96。</returns>
    public static uint GetWindowDpi(IntPtr hwnd)
    {
        if (hwnd == IntPtr.Zero)
        {
            return 96;
        }

        try
        {
            var dpi = GetDpiForWindow(hwnd);
            return dpi == 0 ? 96u : dpi;
        }
        catch
        {
            return 96;
        }
    }

    /// <summary>
    /// 从指定窗口取出 Office/WPS 原生对象模型（绑到该进程，而不是 ROT 里的旧实例）。
    /// </summary>
    /// <param name="hwnd">该进程的窗口。</param>
    /// <returns>IDispatch 对象；失败则为 null。</returns>
    public static object? TryGetNativeOm(IntPtr hwnd)
    {
        if (hwnd == IntPtr.Zero)
        {
            return null;
        }

        var iid = new Guid("00020400-0000-0000-C000-000000000046");
        var hr = AccessibleObjectFromWindow(hwnd, ObjIdNativeOm, ref iid, out var unk);
        return hr == 0 ? unk : null;
    }

    /// <summary>
    /// 在目标屏幕上找与该屏重叠最大的 WPS 窗口（排除已占用句柄）。
    /// </summary>
    /// <param name="bounds">目标屏幕的虚拟桌面矩形。</param>
    /// <param name="exclude">其它路已占用的窗口。</param>
    /// <returns>最匹配的句柄；找不到则为 Zero。</returns>
    public static IntPtr FindWpsWindowOnScreen(
        System.Drawing.Rectangle bounds,
        IEnumerable<IntPtr>? exclude = null,
        int processId = 0)
    {
        var taken = exclude?.Where(item => item != IntPtr.Zero).ToHashSet() ?? [];
        var best = IntPtr.Zero;
        var bestArea = 0L;
        foreach (var window in CaptureTopWindows())
        {
            if (taken.Contains(window.Hwnd) || !IsWpsProcess(window.ProcessName))
            {
                continue;
            }

            if (processId > 0 && GetWindowProcessId(window.Hwnd) != processId)
            {
                continue;
            }

            if (IsEditorTitle(window.Title))
            {
                continue;
            }

            GetWindowRect(window.Hwnd, out var rect);
            var overlap = OverlapArea(bounds, rect);
            if (IsSlideShowTitle(window.Title))
            {
                overlap += 10_000_000L;
            }

            if (overlap > bestArea)
            {
                bestArea = overlap;
                best = window.Hwnd;
            }
        }

        return best;
    }

    /// <summary>
    /// 句柄是否仍是有效窗口。
    /// </summary>
    /// <param name="hwnd">窗口句柄。</param>
    /// <returns>有效则为 true。</returns>
    public static bool IsAlive(IntPtr hwnd)
    {
        return hwnd != IntPtr.Zero && IsWindow(hwnd);
    }

    /// <summary>
    /// 窗口是否主要落在指定屏幕上。
    /// </summary>
    /// <param name="hwnd">窗口。</param>
    /// <param name="bounds">屏幕矩形。</param>
    /// <returns>重叠超过一半窗口面积则为 true。</returns>
    public static bool IsOnScreen(IntPtr hwnd, System.Drawing.Rectangle bounds)
    {
        if (!IsAlive(hwnd))
        {
            return false;
        }

        GetWindowRect(hwnd, out var rect);
        var width = rect.Right - rect.Left;
        var height = rect.Bottom - rect.Top;
        var area = width * (long)height;
        if (area <= 0)
        {
            return false;
        }

        return OverlapArea(bounds, rect) * 2 >= area;
    }

    /// <summary>
    /// 读取显示器的物理像素矩形（含任务栏区域），避免 DPI 换算成主屏逻辑尺寸。
    /// </summary>
    /// <param name="hint">该屏上任意一点所在的矩形。</param>
    /// <returns>rcMonitor 物理像素矩形。</returns>
    public static System.Drawing.Rectangle GetPhysicalMonitorBounds(System.Drawing.Rectangle hint)
    {
        var query = new Rect { Left = hint.Left, Top = hint.Top, Right = hint.Right, Bottom = hint.Bottom };
        var monitor = MonitorFromRect(ref query, MonitorDefaultToNearest);
        var info = new MonitorInfo { Size = Marshal.SizeOf<MonitorInfo>() };
        if (monitor != IntPtr.Zero && GetMonitorInfo(monitor, ref info))
        {
            return System.Drawing.Rectangle.FromLTRB(
                info.Monitor.Left,
                info.Monitor.Top,
                info.Monitor.Right,
                info.Monitor.Bottom);
        }

        return hint;
    }

    /// <summary>
    /// 将窗口铺满指定屏幕的实际像素矩形（不用主屏分辨率），并去掉边框、盖住任务栏。
    /// </summary>
    /// <param name="hwnd">目标窗口。</param>
    /// <param name="bounds">该屏在虚拟桌面上的像素矩形。</param>
    /// <param name="raiseZOrder">true 时置顶盖住任务栏；翻页重铺时应为 false，避免盖住悬浮球。</param>
    public static void FitToBounds(
        IntPtr hwnd,
        System.Drawing.Rectangle bounds,
        bool raiseZOrder = true,
        bool allowEditor = false)
    {
        if (!IsAlive(hwnd) || bounds.Width <= 0 || bounds.Height <= 0)
        {
            return;
        }

        if (!allowEditor && IsEditorTitle(GetTitle(hwnd)))
        {
            return;
        }

        bounds = GetPhysicalMonitorBounds(bounds);
        SetWindowRgn(hwnd, IntPtr.Zero, true);
        var style = GetWindowLong(hwnd, GwlStyle);
        SetWindowLong(hwnd, GwlStyle, style & ~WsCaption & ~WsThickFrame & ~WsBorder & ~WsDlgFrame);

        // NOSENDCHANGING：阻止 WPS 在 WM_WINDOWPOSCHANGING 里把高度钳回主屏 1440
        var flags = SwpShowWindow | SwpFrameChanged | SwpNoCopyBits | SwpNoSendChanging;
        var insertAfter = HwndTopMost;
        if (!raiseZOrder)
        {
            flags |= SwpNoZOrder;
            insertAfter = IntPtr.Zero;
        }

        MoveWindow(hwnd, bounds.X, bounds.Y, bounds.Width, bounds.Height, true);
        SetWindowPos(hwnd, insertAfter, bounds.X, bounds.Y, bounds.Width, bounds.Height, flags);
        FitChildWindows(hwnd, bounds.Width, bounds.Height);
        ShowWindow(hwnd, raiseZOrder ? SwShow : SwShowNoActivate);

        if (!MatchesBounds(hwnd, bounds))
        {
            MoveWindow(hwnd, bounds.X, bounds.Y, bounds.Width, bounds.Height, true);
            SetWindowPos(hwnd, insertAfter, bounds.X, bounds.Y, bounds.Width, bounds.Height, flags);
            FitChildWindows(hwnd, bounds.Width, bounds.Height);
        }
    }

    /// <summary>
    /// 按幻灯片比例覆盖铺满目标屏：窗口可大于屏幕，裁掉多出的边，避免 16:9 在 16:10 上留黑边。
    /// </summary>
    /// <param name="hwnd">放映窗。</param>
    /// <param name="screen">目标屏矩形。</param>
    /// <param name="slideAspect">幻灯片宽高比；无效时按 16:9。</param>
    /// <param name="raiseZOrder">是否置顶。</param>
    public static void FitCoverToMonitor(
        IntPtr hwnd,
        System.Drawing.Rectangle screen,
        double slideAspect,
        bool raiseZOrder = true)
    {
        if (!IsAlive(hwnd))
        {
            return;
        }

        screen = GetPhysicalMonitorBounds(screen);
        var cover = ComputeCoverRect(screen, slideAspect);
        if (MatchesBounds(hwnd, cover, 24))
        {
            ClipWindowToMonitor(hwnd, screen, cover);
            return;
        }

        var style = GetWindowLong(hwnd, GwlStyle);
        SetWindowLong(hwnd, GwlStyle, style & ~WsCaption & ~WsThickFrame & ~WsBorder & ~WsDlgFrame);

        var flags = SwpShowWindow | SwpFrameChanged | SwpNoCopyBits | SwpNoSendChanging;
        var insertAfter = HwndTopMost;
        if (!raiseZOrder)
        {
            flags |= SwpNoZOrder;
            insertAfter = IntPtr.Zero;
        }

        MoveWindow(hwnd, cover.X, cover.Y, cover.Width, cover.Height, true);
        SetWindowPos(hwnd, insertAfter, cover.X, cover.Y, cover.Width, cover.Height, flags);
        FitChildWindows(hwnd, cover.Width, cover.Height);
        ClipWindowToMonitor(hwnd, screen, cover);
        ShowWindow(hwnd, raiseZOrder ? SwShow : SwShowNoActivate);
    }

    /// <summary>
    /// 窗口是否已经按覆盖矩形铺在目标屏上（允许大于屏幕）。
    /// </summary>
    /// <param name="hwnd">窗口。</param>
    /// <param name="screen">目标屏。</param>
    /// <param name="slideAspect">幻灯片宽高比。</param>
    /// <param name="tolerance">像素容差。</param>
    /// <returns>已覆盖则为 true。</returns>
    public static bool MatchesCoverBounds(
        IntPtr hwnd,
        System.Drawing.Rectangle screen,
        double slideAspect,
        int tolerance = 24)
    {
        if (!IsAlive(hwnd))
        {
            return false;
        }

        screen = GetPhysicalMonitorBounds(screen);
        var cover = ComputeCoverRect(screen, slideAspect);
        GetWindowRect(hwnd, out var actual);
        return Math.Abs(actual.Left - cover.Left) <= tolerance &&
               Math.Abs(actual.Top - cover.Top) <= tolerance &&
               Math.Abs(actual.Right - actual.Left - cover.Width) <= tolerance &&
               Math.Abs(actual.Bottom - actual.Top - cover.Height) <= tolerance;
    }

    /// <summary>
    /// 按「覆盖裁切」算出比屏幕更大的窗口矩形，使 16:9 内容能铺满 16:10 屏。
    /// </summary>
    /// <param name="screen">屏幕物理矩形。</param>
    /// <param name="slideAspect">幻灯片宽/高。</param>
    /// <returns>覆盖用窗口矩形。</returns>
    public static System.Drawing.Rectangle ComputeCoverRect(System.Drawing.Rectangle screen, double slideAspect)
    {
        if (slideAspect <= 0.1 || !double.IsFinite(slideAspect))
        {
            slideAspect = 16.0 / 9.0;
        }

        var screenAspect = screen.Height > 0 ? (double)screen.Width / screen.Height : slideAspect;
        int width;
        int height;
        if (slideAspect >= screenAspect)
        {
            height = screen.Height;
            width = (int)Math.Ceiling(screen.Height * slideAspect);
        }
        else
        {
            width = screen.Width;
            height = (int)Math.Ceiling(screen.Width / slideAspect);
        }

        var left = screen.Left + (screen.Width - width) / 2;
        var top = screen.Top + (screen.Height - height) / 2;
        return new System.Drawing.Rectangle(left, top, width, height);
    }

    /// <summary>
    /// 把大于屏幕的放映窗裁到当前屏，避免盖到另一块显示器。
    /// </summary>
    /// <param name="hwnd">放映窗。</param>
    /// <param name="screen">目标屏。</param>
    /// <param name="cover">覆盖窗口矩形。</param>
    private static void ClipWindowToMonitor(
        IntPtr hwnd,
        System.Drawing.Rectangle screen,
        System.Drawing.Rectangle cover)
    {
        var clipLeft = screen.Left - cover.Left;
        var clipTop = screen.Top - cover.Top;
        var region = CreateRectRgn(clipLeft, clipTop, clipLeft + screen.Width, clipTop + screen.Height);
        if (region != IntPtr.Zero)
        {
            SetWindowRgn(hwnd, region, true);
        }
    }

    /// <summary>
    /// 静默置顶，不激活、不重绘闪烁。
    /// </summary>
    /// <param name="hwnd">目标窗口。</param>
    public static void RaiseTopMost(IntPtr hwnd)
    {
        if (!IsAlive(hwnd))
        {
            return;
        }

        SetWindowPos(hwnd, HwndTopMost, 0, 0, 0, 0, SwpNoMove | SwpNoSize | SwpNoActivate);
    }

    /// <summary>
    /// 悬浮条不抢激活，避免被放映窗抢焦时闪一下。
    /// </summary>
    /// <param name="hwnd">悬浮条句柄。</param>
    public static void ApplyFloatBallExStyle(IntPtr hwnd)
    {
        if (!IsAlive(hwnd))
        {
            return;
        }

        var ex = GetWindowLong(hwnd, GwlExStyle);
        SetWindowLong(hwnd, GwlExStyle, ex | WsExNoActivate | WsExToolWindow);
        RaiseTopMost(hwnd);
    }

    /// <summary>
    /// 启用 Win10/11 磨砂玻璃，隐约透出后面的放映画面。
    /// </summary>
    /// <param name="hwnd">悬浮条句柄。</param>
    public static void EnableAcrylicBlur(IntPtr hwnd)
    {
        if (!IsAlive(hwnd))
        {
            return;
        }

        var accent = new AccentPolicy
        {
            AccentState = 4,
            AccentFlags = 2,
            GradientColor = unchecked((int)0x66101820)
        };
        var accentSize = Marshal.SizeOf<AccentPolicy>();
        var accentPtr = Marshal.AllocHGlobal(accentSize);
        try
        {
            Marshal.StructureToPtr(accent, accentPtr, false);
            var data = new WindowCompositionAttributeData
            {
                Attribute = 19,
                Data = accentPtr,
                SizeOfData = accentSize
            };
            SetWindowCompositionAttribute(hwnd, ref data);
        }
        catch
        {
            try
            {
                accent.AccentState = 3;
                Marshal.StructureToPtr(accent, accentPtr, false);
                var data = new WindowCompositionAttributeData
                {
                    Attribute = 19,
                    Data = accentPtr,
                    SizeOfData = accentSize
                };
                SetWindowCompositionAttribute(hwnd, ref data);
            }
            catch
            {
                // 系统不支持模糊时保持半透明底
            }
        }
        finally
        {
            Marshal.FreeHGlobal(accentPtr);
        }
    }

    /// <summary>
    /// 按圆角裁剪窗口，配合磨砂底使用。
    /// </summary>
    /// <param name="hwnd">窗口。</param>
    /// <param name="width">宽度。</param>
    /// <param name="height">高度。</param>
    /// <param name="radius">圆角半径。</param>
    public static void ApplyRoundWindowRegion(IntPtr hwnd, int width, int height, int radius)
    {
        if (!IsAlive(hwnd) || width <= 0 || height <= 0)
        {
            return;
        }

        var region = CreateRoundRectRgn(0, 0, width + 1, height + 1, radius, radius);
        SetWindowRgn(hwnd, region, true);
    }

    /// <summary>
    /// 屏幕坐标处最上层窗口是否属于指定窗口（用于判断悬浮条是否被盖住）。
    /// </summary>
    /// <param name="hwnd">悬浮条。</param>
    /// <param name="screenX">屏幕 X。</param>
    /// <param name="screenY">屏幕 Y。</param>
    /// <returns>该点属于本窗口则为 true。</returns>
    public static bool IsTopAtScreenPoint(IntPtr hwnd, int screenX, int screenY)
    {
        if (!IsAlive(hwnd))
        {
            return false;
        }

        var hit = HitTestFromPoint(new System.Drawing.Point(screenX, screenY));
        return hit == hwnd || IsChild(hwnd, hit);
    }

    /// <summary>
    /// 窗口外框是否已经等于目标屏物理矩形。
    /// </summary>
    /// <param name="hwnd">窗口。</param>
    /// <param name="bounds">目标矩形。</param>
    /// <param name="tolerance">允许的像素误差。</param>
    /// <returns>已对齐则为 true。</returns>
    public static bool MatchesBounds(IntPtr hwnd, System.Drawing.Rectangle bounds, int tolerance = 4)
    {
        if (!IsAlive(hwnd))
        {
            return false;
        }

        bounds = GetPhysicalMonitorBounds(bounds);
        GetWindowRect(hwnd, out var actual);
        return Math.Abs(actual.Left - bounds.Left) <= tolerance &&
               Math.Abs(actual.Top - bounds.Top) <= tolerance &&
               Math.Abs(actual.Right - actual.Left - bounds.Width) <= tolerance &&
               Math.Abs(actual.Bottom - actual.Top - bounds.Height) <= tolerance;
    }

    /// <summary>
    /// 尝试把放映窗提到前台，便于 COM 指令落到该窗口。
    /// </summary>
    /// <param name="hwnd">目标窗口。</param>
    public static void TryFocus(IntPtr hwnd)
    {
        if (hwnd == IntPtr.Zero)
        {
            return;
        }

        ShowWindow(hwnd, SwShow);
        SetForegroundWindow(hwnd);
    }

    /// <summary>
    /// 向窗口投递虚拟键（COM 下一步失败时的兜底）。
    /// </summary>
    /// <param name="hwnd">目标窗口。</param>
    /// <param name="virtualKey">虚拟键码，例如右方向键 0x27。</param>
    public static void PostVirtualKey(IntPtr hwnd, int virtualKey)
    {
        if (hwnd == IntPtr.Zero)
        {
            return;
        }

        PostMessage(hwnd, WmKeyDown, (IntPtr)virtualKey, IntPtr.Zero);
        PostMessage(hwnd, WmKeyUp, (IntPtr)virtualKey, IntPtr.Zero);
    }

    /// <summary>
    /// 判断进程是否为 WPS 演示相关。
    /// </summary>
    /// <param name="processName">进程名。</param>
    /// <returns>是 WPS 则为 true。</returns>
    public static bool IsWpsProcess(string processName)
    {
        return processName is "wps" or "wpp" or "wpsoffice";
    }

    /// <summary>
    /// WPS 365 放映窗是无标题的 Qt 窗，旧版才带 Slide Show 字样。
    /// </summary>
    /// <param name="title">标题。</param>
    /// <param name="className">类名。</param>
    /// <param name="processName">进程名。</param>
    /// <returns>像放映窗则为 true。</returns>
    public static bool IsLikelyShowWindow(string title, string className, string processName)
    {
        if (IsPlaybackHostTitle(title) ||
            IsWebPlayerTitle(title) ||
            IsDocumentSurfaceClass(className) ||
            IsSlideShowTitle(title))
        {
            return true;
        }

        return className.Contains("Qt5QWindow", StringComparison.OrdinalIgnoreCase) &&
               string.IsNullOrWhiteSpace(title);
    }

    /// <summary>
    /// 用 Win32 最小化仍挂在桌面上的 WPS 编辑壳，避免挡住放映。
    /// </summary>
    public static void MinimizeEditorShells()
    {
        foreach (var window in CaptureAllWpsWindows())
        {
            if (!window.Visible || !IsEditorTitle(window.Title))
            {
                continue;
            }

            ShowWindow(window.Hwnd, SwMinimize);
        }
    }

    /// <summary>
    /// 窗口标题是否像旧版 WPS 放映窗。
    /// </summary>
    /// <param name="title">窗口标题。</param>
    /// <returns>是放映窗则为 true。</returns>
    public static bool IsSlideShowTitle(string title)
    {
        return title.Contains("Slide Show", StringComparison.OrdinalIgnoreCase) ||
               title.Contains("放映", StringComparison.OrdinalIgnoreCase) ||
               title.Contains("幻灯片放映", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// 是否为 WPS 编辑窗（不是放映窗）。
    /// </summary>
    /// <param name="title">窗口标题。</param>
    /// <returns>是编辑窗则为 true。</returns>
    public static bool IsEditorTitle(string title)
    {
        if (string.IsNullOrWhiteSpace(title) || IsSlideShowTitle(title))
        {
            return false;
        }

        return title.Contains("WPS Office", StringComparison.OrdinalIgnoreCase) ||
               title.Contains("只读", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// 句柄是否可作为放映窗铺屏（排除编辑壳）。
    /// </summary>
    /// <param name="hwnd">窗口。</param>
    /// <returns>可用则为 true。</returns>
    public static bool IsUsableShowWindow(IntPtr hwnd)
    {
        if (!IsAlive(hwnd))
        {
            return false;
        }

        var className = GetClassName(hwnd);
        if (IsDocumentSurfaceClass(className))
        {
            return true;
        }

        var title = GetTitle(hwnd);
        if (IsSlideShowTitle(title))
        {
            return true;
        }

        return !IsEditorTitle(title) &&
               !className.Contains("Qt5QWindow", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// 读取窗口标题。
    /// </summary>
    /// <param name="hwnd">窗口。</param>
    /// <returns>标题；失败则为空串。</returns>
    public static string GetTitle(IntPtr hwnd)
    {
        if (!IsAlive(hwnd))
        {
            return string.Empty;
        }

        var buffer = new StringBuilder(512);
        GetWindowText(hwnd, buffer, buffer.Capacity);
        return buffer.ToString();
    }

    private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool IsWindow(IntPtr hWnd);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "GetClassName")]
    private static extern int QueryClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll")]
    private static extern bool PrintWindow(IntPtr hwnd, IntPtr hdcBlt, uint nFlags);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr CreateWindowEx(
        int dwExStyle,
        string lpClassName,
        string lpWindowName,
        int dwStyle,
        int x,
        int y,
        int nWidth,
        int nHeight,
        IntPtr hWndParent,
        IntPtr hMenu,
        IntPtr hInstance,
        IntPtr lpParam);

    [DllImport("user32.dll")]
    private static extern bool DestroyWindow(IntPtr hWnd);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr GetModuleHandle(string? lpModuleName);

    [DllImport("user32.dll", EntryPoint = "GetParent")]
    private static extern IntPtr GetAncestorParent(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(IntPtr hWnd, out Rect lpRect);

    [DllImport("user32.dll")]
    private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int cx, int cy, uint uFlags);

    [DllImport("user32.dll")]
    private static extern bool MoveWindow(IntPtr hWnd, int x, int y, int nWidth, int nHeight, bool bRepaint);

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    private static extern uint GetDpiForWindow(IntPtr hwnd);

    [DllImport("user32.dll")]
    private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll")]
    private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool RedrawWindow(IntPtr hWnd, IntPtr lprcUpdate, IntPtr hrgnUpdate, uint flags);

    [DllImport("user32.dll")]
    private static extern IntPtr SetFocus(IntPtr hWnd);

    [DllImport("oleacc.dll")]
    private static extern int AccessibleObjectFromWindow(
        IntPtr hwnd,
        uint dwObjectId,
        ref Guid riid,
        [MarshalAs(UnmanagedType.IUnknown)] out object? ppvObject);

    [DllImport("user32.dll")]
    private static extern IntPtr MonitorFromRect(ref Rect lprc, uint dwFlags);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern bool GetMonitorInfo(IntPtr hMonitor, ref MonitorInfo lpmi);

    [DllImport("user32.dll")]
    private static extern bool GetClientRect(IntPtr hWnd, out Rect lpRect);

    [DllImport("user32.dll")]
    private static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll", EntryPoint = "WindowFromPoint")]
    private static extern IntPtr HitTestFromPoint(System.Drawing.Point point);

    [DllImport("user32.dll")]
    private static extern bool IsChild(IntPtr hWndParent, IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern int SetWindowRgn(IntPtr hWnd, IntPtr hRgn, bool bRedraw);

    [DllImport("gdi32.dll")]
    private static extern IntPtr CreateRoundRectRgn(int x1, int y1, int x2, int y2, int w, int h);

    [DllImport("gdi32.dll")]
    private static extern IntPtr CreateRectRgn(int x1, int y1, int x2, int y2);

    [DllImport("user32.dll")]
    private static extern int SetWindowCompositionAttribute(IntPtr hwnd, ref WindowCompositionAttributeData data);

    [StructLayout(LayoutKind.Sequential)]
    private struct AccentPolicy
    {
        public int AccentState;
        public int AccentFlags;
        public int GradientColor;
        public int AnimationId;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct WindowCompositionAttributeData
    {
        public int Attribute;
        public IntPtr Data;
        public int SizeOfData;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct Rect
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct MonitorInfo
    {
        public int Size;
        public Rect Monitor;
        public Rect Work;
        public uint Flags;
    }

    /// <summary>
    /// 计算屏幕矩形与窗口矩形的重叠面积。
    /// </summary>
    /// <param name="bounds">屏幕矩形。</param>
    /// <param name="rect">窗口矩形。</param>
    /// <returns>重叠像素面积。</returns>
    private static long OverlapArea(System.Drawing.Rectangle bounds, Rect rect)
    {
        var left = Math.Max(bounds.Left, rect.Left);
        var top = Math.Max(bounds.Top, rect.Top);
        var right = Math.Min(bounds.Right, rect.Right);
        var bottom = Math.Min(bounds.Bottom, rect.Bottom);
        var width = right - left;
        var height = bottom - top;
        return width > 0 && height > 0 ? width * (long)height : 0;
    }

    /// <summary>
    /// 把放映窗的子窗口也拉成目标宽高，避免只撑了外壳、内容仍是主屏尺寸。
    /// </summary>
    /// <param name="parent">顶层放映窗。</param>
    /// <param name="width">目标宽度。</param>
    /// <param name="height">目标高度。</param>
    private static void FitChildWindows(IntPtr parent, int width, int height)
    {
        EnumChildWindows(parent, (child, _) =>
        {
            SetWindowPos(child, IntPtr.Zero, 0, 0, width, height, SwpNoZOrder | SwpShowWindow | SwpFrameChanged);
            return true;
        }, IntPtr.Zero);
    }
}
