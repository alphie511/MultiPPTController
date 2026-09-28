using MultiPPTController.Models;

namespace MultiPPTController.Services;

/// <summary>
/// 本机双屏开播冒烟：验证主副屏都出现放映窗且铺到对应物理矩形。
/// </summary>
public static class PlaybackSmoke
{
    /// <summary>
    /// 运行冒烟并写日志。
    /// </summary>
    /// <returns>全部通过为 0，否则为非 0。</returns>
    public static int Run()
    {
        var logPath = Path.Combine(Path.GetTempPath(), "MultiPPT-smoke.log");
        File.WriteAllText(logPath, $"=== smoke {DateTime.Now:O} ==={Environment.NewLine}");
        void Log(string line)
        {
            File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss.fff}] {line}{Environment.NewLine}");
        }

        try
        {
            var screens = new MonitorService().GetScreens();
            Log($"screens={screens.Count}");
            foreach (var screen in screens)
            {
                Log($"  {screen.Label} {screen.Bounds}");
            }

            if (screens.Count < 2)
            {
                Log("FAIL: 需要两块屏幕");
                return 2;
            }

            var deck1 = Path.GetFullPath(@"D:\workstation\MultiPPTController\testdata\deck1.pptx");
            var deck2 = Path.GetFullPath(@"D:\workstation\MultiPPTController\testdata\deck2.pptx");
            if (!File.Exists(deck1) || !File.Exists(deck2))
            {
                Log("FAIL: testdata 缺少 deck1/deck2");
                return 3;
            }

            var primary = screens.First(item => item.IsPrimary);
            var secondary = screens.First(item => !item.IsPrimary);
            var assignments = new List<(DeckItem Deck, DisplayScreen Screen)>
            {
                (new DeckItem(deck1), primary),
                (new DeckItem(deck2), secondary)
            };

            Log("existing processes:");
            foreach (var name in new[] { "wps", "wpp", "wpsoffice", "MultiPPTController" })
            {
                foreach (var process in System.Diagnostics.Process.GetProcessesByName(name))
                {
                    Log($"  pid={process.Id} name={process.ProcessName} title='{process.MainWindowTitle}'");
                }
            }

            using var orchestrator = new PlaybackOrchestrator();
            orchestrator.StartAsync(assignments).GetAwaiter().GetResult();
            Thread.Sleep(1500);

            foreach (var line in orchestrator.DescribeSessionsAsync().GetAwaiter().GetResult())
            {
                Log($"session {line}");
            }

            Log("all wps/wpp/web top windows (incl hidden):");
            foreach (var window in NativeWindowHelper.CaptureAllWpsWindows())
            {
                if (!window.Visible && window.Width * (long)window.Height < 200 * 150 && string.IsNullOrWhiteSpace(window.Title))
                {
                    continue;
                }

                Log($"  pid={window.ProcessId} {window.ProcessName} vis={window.Visible} class={window.ClassName} hwnd=0x{window.Hwnd.ToInt64():X} '{window.Title}' {window.Left},{window.Top} {window.Width}x{window.Height} slideshow={NativeWindowHelper.IsSlideShowTitle(window.Title)} editor={NativeWindowHelper.IsEditorTitle(window.Title)}");
            }

            var shows = NativeWindowHelper.CaptureAllWpsWindows()
                .Where(item => item.Visible &&
                               NativeWindowHelper.IsLikelyShowWindow(item.Title, item.ClassName, item.ProcessName) &&
                               item.Width >= 800 &&
                               item.Height >= 600)
                .ToList();
            foreach (var window in shows)
            {
                var pixels = NativeWindowHelper.SamplePixels(window.Hwnd);
                Log($"show hwnd=0x{window.Hwnd.ToInt64():X} class={window.ClassName} '{window.Title}' {window.Left},{window.Top} {window.Width}x{window.Height} pixels={pixels.Colored}/{pixels.Samples}");
            }

            var primaryOk = shows.Any(item => MatchesScreen(item, primary.Bounds) && NativeWindowHelper.SamplePixels(item.Hwnd).Colored >= 8);
            var secondaryOk = shows.Any(item => MatchesScreen(item, secondary.Bounds) && NativeWindowHelper.SamplePixels(item.Hwnd).Colored >= 8);

            Log($"primaryOk={primaryOk} secondaryOk={secondaryOk} showCount={shows.Count}");

            if (!primaryOk || !secondaryOk || shows.Count < 2)
            {
                orchestrator.StopAsync().GetAwaiter().GetResult();
                Log("FAIL: 主屏或副屏放映窗未就位");
                return 4;
            }

            var beforeLines = orchestrator.ReadProgressLinesAsync().GetAwaiter().GetResult();
            foreach (var line in beforeLines)
            {
                Log($"progress0 {line}");
            }

            var primaryStart = ReadSlide(beforeLines, primary.Label);
            var secondaryStart = ReadSlide(beforeLines, secondary.Label);
            for (var step = 1; step <= 3; step++)
            {
                orchestrator.NextAsync().GetAwaiter().GetResult();
                Thread.Sleep(400);
                var lines = orchestrator.ReadProgressLinesAsync().GetAwaiter().GetResult();
                foreach (var line in lines)
                {
                    Log($"progress{step} {line}");
                }
            }

            var afterLines = orchestrator.ReadProgressLinesAsync().GetAwaiter().GetResult();
            var primaryAfter = ReadSlide(afterLines, primary.Label);
            var secondaryAfter = ReadSlide(afterLines, secondary.Label);
            Log($"advance primary {primaryStart}->{primaryAfter} secondary {secondaryStart}->{secondaryAfter}");
            orchestrator.StopAsync().GetAwaiter().GetResult();
            Log("stopped");

            if (primaryAfter < primaryStart + 2)
            {
                Log("FAIL: 主屏连续三下下一步没有连续翻页");
                return 6;
            }

            if (secondaryAfter > 0 && secondaryAfter < secondaryStart + 1)
            {
                Log("FAIL: 副屏连续下一步没有前进");
                return 7;
            }

            if (secondaryAfter <= 0)
            {
                Log("FAIL: 副屏无法读取翻页进度");
                return 8;
            }

            var cloneAssignments = new List<(DeckItem Deck, DisplayScreen Screen)>
            {
                (new DeckItem(deck1), primary),
                (new DeckItem(deck1), secondary)
            };
            orchestrator.StartAsync(cloneAssignments).GetAwaiter().GetResult();
            Thread.Sleep(1500);
            var cloneShows = NativeWindowHelper.CaptureAllWpsWindows()
                .Where(item => item.Visible &&
                               NativeWindowHelper.IsLikelyShowWindow(item.Title, item.ClassName, item.ProcessName) &&
                               item.Width >= 800 &&
                               item.Height >= 600)
                .ToList();
            foreach (var window in cloneShows)
            {
                Log($"clone show hwnd=0x{window.Hwnd.ToInt64():X} class={window.ClassName} '{window.Title}' {window.Left},{window.Top} {window.Width}x{window.Height}");
            }

            var clonePrimary = cloneShows.Any(item => MatchesScreen(item, primary.Bounds));
            var cloneSecondary = cloneShows.Any(item => MatchesScreen(item, secondary.Bounds));
            Log($"clonePrimaryOk={clonePrimary} cloneSecondaryOk={cloneSecondary} count={cloneShows.Count}");
            orchestrator.StopAsync().GetAwaiter().GetResult();
            Log("clone stopped");

            if (!clonePrimary || !cloneSecondary || cloneShows.Count < 2)
            {
                Log("FAIL: 同一份 PPT 双屏未就位");
                return 5;
            }

            Log("PASS");
            return 0;
        }
        catch (Exception ex)
        {
            Log($"FAIL: {ex}");
            return 1;
        }
        finally
        {
            Console.WriteLine(logPath);
        }
    }

    /// <summary>
    /// 放映窗是否已铺到目标屏。
    /// </summary>
    /// <param name="window">窗口快照。</param>
    /// <param name="bounds">屏幕矩形。</param>
    /// <returns>位置与尺寸吻合则为 true。</returns>
    private static bool MatchesScreen(NativeWindowHelper.ProcessWindow window, System.Drawing.Rectangle bounds)
    {
        return Math.Abs(window.Left - bounds.Left) <= 8 &&
               Math.Abs(window.Top - bounds.Top) <= 8 &&
               Math.Abs(window.Width - bounds.Width) <= 8 &&
               Math.Abs(window.Height - bounds.Height) <= 8;
    }

    /// <summary>
    /// 从进度行里取出指定屏幕的页码。
    /// </summary>
    /// <param name="lines">ReadProgressLines 结果。</param>
    /// <param name="label">屏幕标签。</param>
    /// <returns>页码；读不到则为 0。</returns>
    private static int ReadSlide(IReadOnlyList<string> lines, string label)
    {
        var line = lines.FirstOrDefault(item => item.StartsWith(label, StringComparison.Ordinal));
        if (line is null)
        {
            return 0;
        }

        var start = line.IndexOf('第');
        var end = line.IndexOf('页');
        if (start < 0 || end <= start)
        {
            return 0;
        }

        return int.TryParse(line[(start + 1)..end].Trim(), out var slide) ? slide : 0;
    }
}
