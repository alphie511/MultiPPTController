using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using Microsoft.Win32;

internal static class Program
{
    private static readonly string LogPath = Path.Combine(
        Path.GetTempPath(), "MultiPPT-bind-probe.log");

    [STAThread]
    private static int Main(string[] args)
    {
        if (args.Contains("--dual", StringComparer.OrdinalIgnoreCase))
        {
            return RunDual();
        }

        if (args.Contains("--content", StringComparer.OrdinalIgnoreCase))
        {
            return RunContent();
        }

        if (args.Contains("--twoframes", StringComparer.OrdinalIgnoreCase))
        {
            return RunTwoFrames();
        }
        File.WriteAllText(LogPath, $"=== probe {DateTime.Now:O} ==={Environment.NewLine}");
        Log($"screens={System.Windows.Forms.Screen.AllScreens.Length}");
        foreach (var screen in System.Windows.Forms.Screen.AllScreens)
        {
            Log($"  primary={screen.Primary} {screen.DeviceName} {screen.Bounds}");
        }

        Log("existing wps processes:");
        foreach (var process in Snapshot())
        {
            Log($"  pid={process.Id} name={process.ProcessName} cmd={ReadCommandLine(process.Id)}");
        }

        var exe = ResolveWpsExecutable();
        Log($"wps exe={exe}");
        if (string.IsNullOrWhiteSpace(exe) || !File.Exists(exe))
        {
            Log("FAIL: wps exe missing");
            Console.WriteLine(LogPath);
            return 2;
        }

        var before = Snapshot().Select(item => item.Id).ToHashSet();
        Process? started = null;
        try
        {
            started = Process.Start(new ProcessStartInfo
            {
                FileName = exe,
                Arguments = "/wpp /Automation",
                UseShellExecute = false
            });
            Log($"Process.Start pid={started?.Id} hasExited={started?.HasExited}");
        }
        catch (Exception ex)
        {
            Log($"Process.Start failed: {ex}");
        }

        Thread.Sleep(2500);
        Log("after start:");
        foreach (var process in Snapshot())
        {
            var mark = before.Contains(process.Id) ? "old" : "NEW";
            Log($"  [{mark}] pid={process.Id} name={process.ProcessName} title={SafeTitle(process)} cmd={ReadCommandLine(process.Id)}");
            foreach (var hwnd in EnumerateTopWindows(process.Id))
            {
                Log($"    hwnd=0x{hwnd.ToInt64():X} visible={IsWindowVisible(hwnd)} rect={RectText(hwnd)} title={WindowTitle(hwnd)}");
            }
        }

        foreach (var process in Snapshot().Where(item => !before.Contains(item.Id)))
        {
            Log($"NativeOM pid={process.Id} name={process.ProcessName} => {DescribeBind(TryNativeOm(process.Id))}");
            Log($"ROT pid={process.Id} => {DescribeBind(TryRot(process.Id))}");
        }

        var created = TryCreate();
        Log($"CreateInstance={DescribeBind(created)}");
        if (created is not null)
        {
            dynamic app = created;
            try
            {
                app.Visible = true;
                app.DisplayAlerts = 0;
            }
            catch (Exception ex)
            {
                Log($"configure ex={ex.Message}");
            }

            Log($"CreateInstance after Visible={DescribeBind(created)}");
            try
            {
                foreach (dynamic window in app.Windows)
                {
                    try
                    {
                        var hwnd = new IntPtr(Convert.ToInt64(window.HWND));
                        Log($"  app.Window hwnd=0x{hwnd.ToInt64():X} pid={GetPid(hwnd)}");
                    }
                    catch (Exception ex)
                    {
                        Log($"  app.Window ex={ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"app.Windows ex={ex.Message}");
            }

            var deck = @"D:\workstation\MultiPPTController\testdata\deck1.pptx";
            if (File.Exists(deck))
            {
                try
                {
                    var pres = app.Presentations.Open(deck, -1, 0, -1);
                    Log("opened deck1");
                    var settings = pres.SlideShowSettings;
                    settings.ShowType = 2;
                    var show = settings.Run();
                    Thread.Sleep(400);
                    try
                    {
                        var hwnd = new IntPtr(Convert.ToInt64(show.HWND));
                        Log($"show hwnd=0x{hwnd.ToInt64():X} pid={GetPid(hwnd)} rect={RectText(hwnd)}");
                    }
                    catch (Exception ex)
                    {
                        Log($"show hwnd ex={ex.Message}");
                    }

                    try
                    {
                        show.View.Exit();
                    }
                    catch
                    {
                        // ignore
                    }

                    try
                    {
                        pres.Close();
                    }
                    catch
                    {
                        // ignore
                    }
                }
                catch (Exception ex)
                {
                    Log($"open/run ex={ex.GetType().Name}: {ex.Message}");
                }
            }
        }

        Console.WriteLine(LogPath);
        return 0;
    }

    /// <summary>
    /// 探测同一 KWPP 打开两份文稿时能否拆出两扇 PP12 文档框。
    /// </summary>
    /// <returns>完成码。</returns>
    private static int RunTwoFrames()
    {
        File.WriteAllText(LogPath, $"=== twoframes {DateTime.Now:O} ==={Environment.NewLine}");
        var type = Type.GetTypeFromProgID("KWPP.Application", throwOnError: false);
        dynamic app = Activator.CreateInstance(type!)!;
        try { app.Visible = true; app.DisplayAlerts = 0; } catch { /* ignore */ }

        var deck1 = @"D:\workstation\MultiPPTController\testdata\deck1.pptx";
        var deck2 = @"D:\workstation\MultiPPTController\testdata\deck2.pptx";
        dynamic pres1 = app.Presentations.Open(deck1, -1, 0, -1);
        Log("opened deck1");
        DumpAppWindows(app, "after-deck1");
        DumpInteresting("after-deck1");

        try
        {
            var extra = pres1.NewWindow();
            Log($"NewWindow={Safe(() => Convert.ToString(extra))}");
        }
        catch (Exception ex)
        {
            Log($"NewWindow ex={ex.Message}");
        }

        DumpAppWindows(app, "after-newwindow");
        DumpInteresting("after-newwindow");

        dynamic pres2 = app.Presentations.Open(deck2, -1, 0, -1);
        Log("opened deck2");
        DumpAppWindows(app, "after-deck2");
        DumpInteresting("after-deck2");
        DumpAllPp12("after-deck2");

        try
        {
            pres1.Windows[1].Activate();
            Log("activated pres1.Windows[1]");
        }
        catch (Exception ex)
        {
            Log($"activate pres1 ex={ex.Message}");
        }

        try
        {
            Log($"pres1.Windows.Count={Safe(() => Convert.ToString(pres1.Windows.Count))} pres2.Windows.Count={Safe(() => Convert.ToString(pres2.Windows.Count))} app.Windows.Count={Safe(() => Convert.ToString(app.Windows.Count))}");
        }
        catch (Exception ex)
        {
            Log($"counts ex={ex.Message}");
        }

        try { pres2.Close(); } catch { /* ignore */ }
        try { pres1.Close(); } catch { /* ignore */ }
        try { app.Quit(); } catch { /* ignore */ }
        Console.WriteLine(LogPath);
        return 0;
    }

    private static void DumpAllPp12(string tag)
    {
        Log($"all PP12 {tag}:");
        foreach (var root in EnumerateAllWps())
        {
            if (!ClassNameOf(root).Equals("PP12FrameClass", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            GetWindowRect(root, out var rect);
            Log($"  root vis={IsWindowVisible(root)} hwnd=0x{root.ToInt64():X} '{WindowTitle(root)}' {rect.Left},{rect.Top} {rect.Right - rect.Left}x{rect.Bottom - rect.Top}");
            EnumChildWindows(root, (child, _) =>
            {
                if (!ClassNameOf(child).Equals("PP12FrameClass", StringComparison.OrdinalIgnoreCase) &&
                    !ClassNameOf(child).Equals("mdiClass", StringComparison.OrdinalIgnoreCase) &&
                    !WindowTitle(child).Contains("pptx", StringComparison.OrdinalIgnoreCase) &&
                    !WindowTitle(child).Contains("Reading", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }

                GetWindowRect(child, out var childRect);
                Log($"    vis={IsWindowVisible(child)} class={ClassNameOf(child)} hwnd=0x{child.ToInt64():X} '{WindowTitle(child)}' {childRect.Left},{childRect.Top} {childRect.Right - childRect.Left}x{childRect.Bottom - childRect.Top}");
                return true;
            }, IntPtr.Zero);
        }
    }

    private static void DumpAppWindows(dynamic app, string tag)
    {
        try
        {
            Log($"app.Windows {tag} count={Safe(() => Convert.ToString(app.Windows.Count))}");
            foreach (dynamic window in app.Windows)
            {
                Log($"  caption={Safe(() => Convert.ToString(window.Caption))} hwnd={Safe(() => Convert.ToString(window.HWND))} state={Safe(() => Convert.ToString(window.WindowState))}");
            }
        }
        catch (Exception ex)
        {
            Log($"app.Windows ex={ex.Message}");
        }
    }

    /// <summary>
    /// 诊断窗口放映为何全屏空白：对比 ShowType、揭窗、铺子窗前后的像素与子窗口。
    /// </summary>
    /// <returns>完成码。</returns>
    private static int RunContent()
    {
        File.WriteAllText(LogPath, $"=== content {DateTime.Now:O} ==={Environment.NewLine}");
        var type = Type.GetTypeFromProgID("KWPP.Application", throwOnError: false);
        if (type is null)
        {
            Log("FAIL no KWPP");
            return 2;
        }

        dynamic app = Activator.CreateInstance(type)!;
        try
        {
            app.Visible = true;
            app.DisplayAlerts = 0;
        }
        catch
        {
            // ignore
        }

        var deck = @"D:\workstation\MultiPPTController\testdata\deck1.pptx";
        var primary = System.Windows.Forms.Screen.PrimaryScreen!.Bounds;
        DumpProcesses("before-open");

        dynamic pres = app.Presentations.Open(deck, -1, 0, -1);
        Log($"opened slides={Safe(() => Convert.ToString(pres.Slides.Count))}");

        foreach (var showType in new[] { 2, 1 })
        {
            Log($"----- ShowType={showType} Run -----");
            dynamic settings = pres.SlideShowSettings;
            settings.ShowType = showType;
            settings.ShowWithAnimation = -1;
            var show = settings.Run();
            Thread.Sleep(800);
            DumpComShow(app, show, "after-run");
            DumpInteresting("after-run");

            var hidden = FindHiddenQt();
            Log($"hiddenQt={hidden.Count}");
            foreach (var hwnd in hidden)
            {
                ShowWindow(hwnd, 9);
                ShowWindow(hwnd, 8);
                Log($"  revealed 0x{hwnd.ToInt64():X} {RectText(hwnd)} pixels={PixelStats(hwnd)}");
            }

            Thread.Sleep(300);
            DumpInteresting("after-reveal");

            var qt = FindVisibleQt();
            if (qt != IntPtr.Zero)
            {
                MoveWindow(qt, primary.X, primary.Y, primary.Width, primary.Height, true);
                Thread.Sleep(200);
                Log($"after-move 0x{qt.ToInt64():X} {RectText(qt)} pixels={PixelStats(qt)}");
                DumpChildren(qt, "after-move");

                EnumChildWindows(qt, (child, _) =>
                {
                    SetWindowPos(child, IntPtr.Zero, 0, 0, primary.Width, primary.Height, 0x0004 | 0x0040 | 0x0020);
                    return true;
                }, IntPtr.Zero);
                Thread.Sleep(200);
                Log($"after-fitchild 0x{qt.ToInt64():X} {RectText(qt)} pixels={PixelStats(qt)}");
                DumpChildren(qt, "after-fitchild");
            }

            try
            {
                show.View.GotoClick(0);
            }
            catch (Exception ex)
            {
                Log($"GotoClick ex={ex.Message}");
            }

            try
            {
                Log($"view.pos={Safe(() => Convert.ToString(show.View.CurrentShowPosition))} slide={Safe(() => Convert.ToString(show.View.Slide.SlideIndex))}");
                show.View.Next();
                Thread.Sleep(300);
                Log($"after-next pos={Safe(() => Convert.ToString(show.View.CurrentShowPosition))}");
            }
            catch (Exception ex)
            {
                Log($"view next ex={ex.Message}");
            }

            DumpInteresting("after-next");

            try
            {
                show.View.Exit();
            }
            catch
            {
                // ignore
            }

            Thread.Sleep(400);
        }

        try
        {
            pres.Close();
        }
        catch
        {
            // ignore
        }

        try
        {
            app.Quit();
        }
        catch
        {
            // ignore
        }

        Console.WriteLine(LogPath);
        return 0;
    }

    private static void DumpComShow(dynamic app, dynamic show, string tag)
    {
        Log($"COM {tag} hwnd={Safe(() => Convert.ToString(show.HWND))} left={Safe(() => Convert.ToString(show.Left))} top={Safe(() => Convert.ToString(show.Top))} w={Safe(() => Convert.ToString(show.Width))} h={Safe(() => Convert.ToString(show.Height))} pos={Safe(() => Convert.ToString(show.View.CurrentShowPosition))} count={Safe(() => Convert.ToString(app.SlideShowWindows.Count))}");
    }

    private static void DumpProcesses(string tag)
    {
        Log($"processes {tag}:");
        foreach (var name in new[] { "wps", "wpp", "wpsoffice" })
        {
            foreach (var process in Process.GetProcessesByName(name))
            {
                Log($"  pid={process.Id} {process.ProcessName} '{SafeTitle(process)}'");
            }
        }
    }

    private static void DumpInteresting(string tag)
    {
        Log($"windows {tag}:");
        foreach (var hwnd in EnumerateAllWps())
        {
            GetWindowRect(hwnd, out var rect);
            var width = rect.Right - rect.Left;
            var height = rect.Bottom - rect.Top;
            var visible = IsWindowVisible(hwnd);
            var title = WindowTitle(hwnd);
            var className = ClassNameOf(hwnd);
            if (!visible && width * (long)height < 200 * 150 && string.IsNullOrWhiteSpace(title))
            {
                continue;
            }

            GetWindowThreadProcessId(hwnd, out var pid);
            Log($"  pid={pid} vis={visible} class={className} hwnd=0x{hwnd.ToInt64():X} '{title}' {rect.Left},{rect.Top} {width}x{height} pixels={PixelStats(hwnd)}");
            if (visible && width >= 400 && height >= 280)
            {
                DumpChildren(hwnd, tag);
            }
        }
    }

    private static void DumpChildren(IntPtr root, string tag)
    {
        var n = 0;
        EnumChildWindows(root, (child, _) =>
        {
            n++;
            if (n <= 12)
            {
                GetWindowRect(child, out var rect);
                Log($"    child vis={IsWindowVisible(child)} class={ClassNameOf(child)} hwnd=0x{child.ToInt64():X} '{WindowTitle(child)}' {rect.Left},{rect.Top} {rect.Right - rect.Left}x{rect.Bottom - rect.Top}");
            }

            return true;
        }, IntPtr.Zero);
        Log($"    childCount={n} tag={tag}");
    }

    private static List<IntPtr> FindHiddenQt()
    {
        var list = new List<IntPtr>();
        foreach (var hwnd in EnumerateAllWps())
        {
            if (IsWindowVisible(hwnd) || !ClassNameOf(hwnd).Contains("Qt5QWindow", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            GetWindowRect(hwnd, out var rect);
            if (rect.Right - rect.Left >= 400 && rect.Bottom - rect.Top >= 280)
            {
                list.Add(hwnd);
            }
        }

        return list;
    }

    private static IntPtr FindVisibleQt()
    {
        foreach (var hwnd in EnumerateAllWps())
        {
            if (!IsWindowVisible(hwnd) || !ClassNameOf(hwnd).Contains("Qt5QWindow", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            GetWindowRect(hwnd, out var rect);
            if (rect.Right - rect.Left >= 400 && rect.Bottom - rect.Top >= 280)
            {
                return hwnd;
            }
        }

        return IntPtr.Zero;
    }

    private static IReadOnlyList<IntPtr> EnumerateAllWps()
    {
        var list = new List<IntPtr>();
        EnumWindows((hwnd, _) =>
        {
            GetWindowThreadProcessId(hwnd, out var pid);
            try
            {
                var name = Process.GetProcessById((int)pid).ProcessName.ToLowerInvariant();
                if (name is "wps" or "wpp" or "wpsoffice")
                {
                    list.Add(hwnd);
                }
            }
            catch
            {
                // ignore
            }

            return true;
        }, IntPtr.Zero);
        return list;
    }

    private static string PixelStats(IntPtr hwnd)
    {
        GetWindowRect(hwnd, out var rect);
        var width = Math.Max(0, rect.Right - rect.Left);
        var height = Math.Max(0, rect.Bottom - rect.Top);
        if (width < 8 || height < 8 || width > 4000 || height > 3000)
        {
            return "skip";
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
            for (var y = 8; y < height; y += Math.Max(8, height / 20))
            {
                for (var x = 8; x < width; x += Math.Max(8, width / 20))
                {
                    var color = bmp.GetPixel(x, y);
                    samples++;
                    if (color.R + color.G + color.B > 40 && !(color.R > 180 && color.G > 180 && color.B > 180 && Math.Abs(color.R - color.G) < 12))
                    {
                        colored++;
                    }
                }
            }

            return $"colored={colored}/{samples}";
        }
        catch (Exception ex)
        {
            return $"err={ex.Message}";
        }
    }

    private static string ClassNameOf(IntPtr hwnd)
    {
        var buffer = new StringBuilder(256);
        GetClassName(hwnd, buffer, buffer.Capacity);
        return buffer.ToString();
    }

    private static string Safe(Func<string?> read)
    {
        try
        {
            return read() ?? "";
        }
        catch (Exception ex)
        {
            return $"ex:{ex.Message}";
        }
    }

    /// <summary>
    /// 验证同一 KWPP 能否同时开两路窗口放映，并铺到主副屏。
    /// </summary>
    /// <returns>成功为 0。</returns>
    private static int RunDual()
    {
        File.WriteAllText(LogPath, $"=== dual {DateTime.Now:O} ==={Environment.NewLine}");
        var type = Type.GetTypeFromProgID("KWPP.Application", throwOnError: false);
        if (type is null)
        {
            Log("FAIL no KWPP");
            return 2;
        }

        dynamic app = Activator.CreateInstance(type)!;
        try
        {
            app.Visible = true;
            app.DisplayAlerts = 0;
        }
        catch
        {
            // ignore
        }

        var screens = System.Windows.Forms.Screen.AllScreens
            .OrderByDescending(item => item.Primary)
            .ToArray();
        var decks = new[]
        {
            @"D:\workstation\MultiPPTController\testdata\deck1.pptx",
            @"D:\workstation\MultiPPTController\testdata\deck2.pptx"
        };
        var shows = new List<(object Show, IntPtr Hwnd, System.Drawing.Rectangle Bounds)>();

        for (var i = 0; i < Math.Min(screens.Length, decks.Length); i++)
        {
            var before = CaptureVisible();
            dynamic pres = app.Presentations.Open(decks[i], -1, 0, -1);
            dynamic settings = pres.SlideShowSettings;
            settings.ShowType = 2;
            settings.ShowWithAnimation = -1;
            var show = settings.Run();
            Thread.Sleep(500);
            var after = CaptureVisible();
            var created = after
                .Where(item => before.All(old => old.Hwnd != item.Hwnd))
                .OrderByDescending(item => item.Area)
                .ToList();
            Log($"deck{i + 1} new windows: {string.Join(" | ", created.Select(item => $"{item.Name} pid={item.Pid} {item.Rect}"))}");
            var hwnd = created.FirstOrDefault().Hwnd;
            if (hwnd != IntPtr.Zero)
            {
                Fit(hwnd, screens[i].Bounds);
                Thread.Sleep(200);
                GetWindowRect(hwnd, out var rect);
                Log($"deck{i + 1} after fit hwnd=0x{hwnd.ToInt64():X} rect={rect.Left},{rect.Top} {rect.Right - rect.Left}x{rect.Bottom - rect.Top} target={screens[i].Bounds}");
            }

            shows.Add((show, hwnd, screens[i].Bounds));
        }

        Log("visible after both runs:");
        foreach (var window in CaptureVisible())
        {
            Log($"  {window.Name} pid={window.Pid} hwnd=0x{window.Hwnd.ToInt64():X} {window.Rect} {window.Title}");
        }

        Thread.Sleep(2500);
        foreach (var item in shows)
        {
            try
            {
                dynamic show = item.Show;
                show.View.Exit();
            }
            catch
            {
                // ignore
            }
        }

        try
        {
            app.Quit();
        }
        catch
        {
            // ignore
        }

        Console.WriteLine(LogPath);
        return 0;
    }

    private readonly record struct VisWin(IntPtr Hwnd, int Pid, string Name, string Title, string Rect, long Area);

    private static List<VisWin> CaptureVisible()
    {
        var list = new List<VisWin>();
        EnumWindows((hwnd, _) =>
        {
            if (!IsWindowVisible(hwnd))
            {
                return true;
            }

            GetWindowRect(hwnd, out var rect);
            var width = rect.Right - rect.Left;
            var height = rect.Bottom - rect.Top;
            if (width < 200 || height < 150)
            {
                return true;
            }

            GetWindowThreadProcessId(hwnd, out var pid);
            string name;
            try
            {
                name = Process.GetProcessById((int)pid).ProcessName.ToLowerInvariant();
            }
            catch
            {
                return true;
            }

            if (name is not ("wps" or "wpp" or "wpsoffice"))
            {
                return true;
            }

            list.Add(new VisWin(hwnd, (int)pid, name, WindowTitle(hwnd), $"{rect.Left},{rect.Top} {width}x{height}", width * (long)height));
            return true;
        }, IntPtr.Zero);
        return list;
    }

    private static void Fit(IntPtr hwnd, System.Drawing.Rectangle bounds)
    {
        MoveWindow(hwnd, bounds.X, bounds.Y, bounds.Width, bounds.Height, true);
        SetWindowPos(hwnd, new IntPtr(-1), bounds.X, bounds.Y, bounds.Width, bounds.Height, 0x0040 | 0x0020 | 0x0100 | 0x0400);
    }

    private static void Log(string line)
    {
        var text = $"[{DateTime.Now:HH:mm:ss.fff}] {line}";
        Console.WriteLine(text);
        File.AppendAllText(LogPath, text + Environment.NewLine);
    }

    private static string DescribeBind(object? app)
    {
        if (app is null)
        {
            return "null";
        }

        try
        {
            dynamic dyn = app;
            var name = "";
            try
            {
                name = Convert.ToString(dyn.Name) ?? "";
            }
            catch
            {
                // ignore
            }

            var hwnd = IntPtr.Zero;
            try
            {
                hwnd = new IntPtr(Convert.ToInt64(dyn.HWND));
            }
            catch
            {
                // ignore
            }

            var pid = hwnd == IntPtr.Zero ? 0 : GetPid(hwnd);
            return $"ok name={name} hwnd=0x{hwnd.ToInt64():X} pid={pid}";
        }
        catch (Exception ex)
        {
            return $"error {ex.GetType().Name}: {ex.Message}";
        }
    }

    private static object? TryGetActive()
    {
        return TryRot(0);
    }

    private static object? TryCreate()
    {
        try
        {
            var type = Type.GetTypeFromProgID("KWPP.Application", throwOnError: false);
            return type is null ? null : Activator.CreateInstance(type);
        }
        catch (Exception ex)
        {
            Log($"  CreateInstance ex={ex.GetType().Name}: {ex.Message}");
            return null;
        }
    }

    private static object? TryNativeOm(int processId)
    {
        if (processId <= 0)
        {
            return null;
        }

        foreach (var hwnd in EnumerateTopWindows(processId))
        {
            foreach (var node in CollectTree(hwnd))
            {
                var raw = TryGetNativeOm(node);
                if (raw is null)
                {
                    continue;
                }

                dynamic om = raw;
                try
                {
                    var app = om.Application;
                    if (app is not null)
                    {
                        _ = app.Presentations;
                        return app;
                    }
                }
                catch
                {
                    // next
                }

                try
                {
                    _ = om.Presentations;
                    return om;
                }
                catch
                {
                    // next
                }
            }
        }

        return null;
    }

    private static object? TryRot(int processId)
    {
        if (GetRunningObjectTable(0, out var rot) != 0 || rot is null)
        {
            return null;
        }

        try
        {
            rot.EnumRunning(out var enumerator);
            if (enumerator is null)
            {
                return null;
            }

            var monikers = new IMoniker[1];
            while (enumerator.Next(1, monikers, IntPtr.Zero) == 0)
            {
                var moniker = monikers[0];
                try
                {
                    rot.GetObject(moniker, out var obj);
                    if (obj is null)
                    {
                        continue;
                    }

                    try
                    {
                        dynamic app = obj;
                        _ = app.Presentations;
                        if (processId <= 0)
                        {
                            return obj;
                        }

                        var hwnd = new IntPtr(Convert.ToInt64(app.HWND));
                        if (GetPid(hwnd) == processId)
                        {
                            return obj;
                        }
                    }
                    catch
                    {
                        // not kwpp
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(moniker);
                }
            }
        }
        catch (Exception ex)
        {
            Log($"  ROT ex={ex.GetType().Name}: {ex.Message}");
        }
        finally
        {
            Marshal.ReleaseComObject(rot);
        }

        return null;
    }

    private static IEnumerable<Process> Snapshot()
    {
        foreach (var name in new[] { "wps", "wpp", "wpsoffice" })
        {
            foreach (var process in Process.GetProcessesByName(name))
            {
                yield return process;
            }
        }
    }

    private static string? ResolveWpsExecutable()
    {
        var clsid = Registry.ClassesRoot.OpenSubKey(@"KWPP.Application\CLSID")?.GetValue(null) as string;
        if (string.IsNullOrWhiteSpace(clsid))
        {
            return null;
        }

        var localServer = Registry.ClassesRoot.OpenSubKey($@"CLSID\{clsid}\LocalServer32")?.GetValue(null) as string
                          ?? Registry.ClassesRoot.OpenSubKey($@"WOW6432Node\CLSID\{clsid}\LocalServer32")?.GetValue(null) as string;
        if (string.IsNullOrWhiteSpace(localServer))
        {
            return null;
        }

        var text = localServer.Trim();
        if (text.StartsWith('"'))
        {
            var end = text.IndexOf('"', 1);
            if (end > 1)
            {
                return text[1..end];
            }
        }

        return text.Split(' ', 2, StringSplitOptions.RemoveEmptyEntries)[0].Trim('"');
    }

    private static string ReadCommandLine(int pid)
    {
        try
        {
            using var searcher = new System.Management.ManagementObjectSearcher(
                $"SELECT CommandLine FROM Win32_Process WHERE ProcessId={pid}");
            foreach (var item in searcher.Get())
            {
                return Convert.ToString(item["CommandLine"]) ?? "";
            }
        }
        catch
        {
            // ignore
        }

        return "";
    }

    private static string SafeTitle(Process process)
    {
        try
        {
            return process.MainWindowTitle;
        }
        catch
        {
            return "";
        }
    }

    private static IReadOnlyList<IntPtr> EnumerateTopWindows(int processId)
    {
        var list = new List<IntPtr>();
        EnumWindows((hwnd, _) =>
        {
            GetWindowThreadProcessId(hwnd, out var pid);
            if ((int)pid == processId)
            {
                list.Add(hwnd);
            }

            return true;
        }, IntPtr.Zero);
        return list;
    }

    private static IReadOnlyList<IntPtr> CollectTree(IntPtr root)
    {
        var list = new List<IntPtr> { root };
        EnumChildWindows(root, (child, _) =>
        {
            list.Add(child);
            return true;
        }, IntPtr.Zero);
        return list;
    }

    private static object? TryGetNativeOm(IntPtr hwnd)
    {
        var iid = new Guid("00020400-0000-0000-C000-000000000046");
        var hr = AccessibleObjectFromWindow(hwnd, 0xFFFFFFF0, ref iid, out var unk);
        return hr == 0 ? unk : null;
    }

    private static int GetPid(IntPtr hwnd)
    {
        GetWindowThreadProcessId(hwnd, out var pid);
        return (int)pid;
    }

    private static string WindowTitle(IntPtr hwnd)
    {
        var buffer = new StringBuilder(512);
        GetWindowText(hwnd, buffer, buffer.Capacity);
        return buffer.ToString();
    }

    private static string RectText(IntPtr hwnd)
    {
        GetWindowRect(hwnd, out var rect);
        return $"{rect.Left},{rect.Top} {rect.Right - rect.Left}x{rect.Bottom - rect.Top}";
    }

    private delegate bool EnumProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumChildWindows(IntPtr hWndParent, EnumProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(IntPtr hWnd, out Rect lpRect);

    [DllImport("user32.dll")]
    private static extern bool MoveWindow(IntPtr hWnd, int x, int y, int nWidth, int nHeight, bool bRepaint);

    [DllImport("user32.dll")]
    private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int cx, int cy, uint uFlags);

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll")]
    private static extern bool PrintWindow(IntPtr hwnd, IntPtr hdcBlt, uint nFlags);

    [DllImport("oleacc.dll")]
    private static extern int AccessibleObjectFromWindow(
        IntPtr hwnd,
        uint dwObjectId,
        ref Guid riid,
        [MarshalAs(UnmanagedType.IUnknown)] out object? ppvObject);

    [DllImport("ole32.dll")]
    private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable pprot);

    [StructLayout(LayoutKind.Sequential)]
    private struct Rect
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }
}
