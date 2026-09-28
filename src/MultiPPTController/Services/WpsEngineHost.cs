using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Microsoft.Win32;
using MultiPPTController.Models;

namespace MultiPPTController.Services;

/// <summary>
/// 本机 WPS 演示引擎（方案 B 只服务非主屏：一份 KWPP、一路文稿）。
/// </summary>
public sealed class WpsEngineHost
{
    private const int MsoFalse = 0;
    private readonly List<WpsPlaybackSession> _sessions = [];

    /// <summary>
    /// 绑定已启动的进程与 COM 对象。
    /// </summary>
    /// <param name="application">该进程的 KWPP.Application。</param>
    /// <param name="process">WPS 进程；无法关联时为 null。</param>
    public WpsEngineHost(dynamic application, Process? process)
    {
        Application = application;
        Process = process;
        int comPid = TryReadProcessId((object)application);
        ProcessId = comPid > 0 ? comPid : process?.Id ?? 0;
    }

    /// <summary>该路 WPS Application。</summary>
    public dynamic Application { get; }

    /// <summary>对应操作系统进程。</summary>
    public Process? Process { get; }

    /// <summary>进程 ID；未知则为 0。</summary>
    public int ProcessId { get; }

    /// <summary>该引擎上的全部放映会话。</summary>
    public IReadOnlyList<WpsPlaybackSession> Sessions => _sessions;

    /// <summary>
    /// 通过 COM 启动或连接 WPS 演示（真正的引擎是 wpp.exe，不是 wps.exe 启动器）。
    /// </summary>
    /// <returns>已绑定的引擎宿主。</returns>
    public static WpsEngineHost Launch()
    {
        var before = SnapshotEngineProcessIds();
        var application = CreateKwppApplication();
        Configure(application);
        var process = WaitForEngineProcess(before);
        return new WpsEngineHost(application, process);
    }

    /// <summary>
    /// 在该进程中打开文稿并开始副屏放映（方案 B 只应开一路）。
    /// </summary>
    /// <param name="deck">片库条目。</param>
    /// <param name="screen">目标屏幕。</param>
    /// <param name="openPath">打开路径。</param>
    /// <param name="ownsTempCopy">是否删除临时副本。</param>
    public void StartSession(DeckItem deck, DisplayScreen screen, string openPath, bool ownsTempCopy)
    {
        var session = new WpsPlaybackSession(deck, screen, openPath, ownsTempCopy)
        {
            EngineProcessId = ProcessId
        };
        session.Open(Application);
        session.StartShow(_sessions.Select(item => item.ShowHwnd));
        _sessions.Add(session);
        NativeWindowHelper.HideEmptyEngineOverlays();
    }

    /// <summary>
    /// 结束放映并退出该 WPS 进程。
    /// </summary>
    public void Quit()
    {
        foreach (var session in _sessions.ToList())
        {
            try
            {
                session.Close();
            }
            catch
            {
                // 继续退出进程
            }
        }

        _sessions.Clear();
        try
        {
            Application.Quit();
        }
        catch
        {
            // 用户可能已关闭
        }

        try
        {
            Marshal.FinalReleaseComObject(Application);
        }
        catch
        {
            // 忽略
        }

        try
        {
            if (Process is { HasExited: false })
            {
                if (!Process.WaitForExit(2000))
                {
                    Process.Kill(entireProcessTree: true);
                }
            }
        }
        catch
        {
            // 退出阶段忽略
        }
    }

    /// <summary>
    /// 收起该进程里的编辑窗口，避免挡住放映。
    /// </summary>
    public void MinimizeEditorWindows()
    {
        var showHwnds = _sessions.Select(item => item.ShowHwnd).ToHashSet();
        try
        {
            foreach (dynamic window in Application.Windows)
            {
                try
                {
                    var hwnd = IntPtr.Zero;
                    try
                    {
                        hwnd = new IntPtr(Convert.ToInt64(window.HWND));
                    }
                    catch
                    {
                        // 无 HWND
                    }

                    if (hwnd != IntPtr.Zero && showHwnds.Contains(hwnd))
                    {
                        continue;
                    }

                    var title = hwnd == IntPtr.Zero ? string.Empty : NativeWindowHelper.GetTitle(hwnd);
                    if (!NativeWindowHelper.IsEditorTitle(title))
                    {
                        continue;
                    }

                    window.WindowState = 2;
                }
                catch
                {
                    // 个别窗口不可最小化
                }
            }
        }
        catch
        {
            // 版本不支持
        }

        NativeWindowHelper.MinimizeEditorShells();
    }

    /// <summary>
    /// 从注册表解析 WPS 演示可执行文件并带 /Automation 启动。
    /// </summary>
    /// <returns>已启动进程；失败则为 null。</returns>
    private static Process? TryStartAutomationProcess()
    {
        var exe = ResolveWpsExecutable();
        if (string.IsNullOrWhiteSpace(exe) || !File.Exists(exe))
        {
            return null;
        }

        try
        {
            return Process.Start(new ProcessStartInfo
            {
                FileName = exe,
                Arguments = "/wpp /Automation",
                UseShellExecute = false
            });
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// 解析 KWPP.Application 对应的 wps.exe 路径。
    /// </summary>
    /// <returns>绝对路径；找不到则为 null。</returns>
    private static string? ResolveWpsExecutable()
    {
        try
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

            var parts = text.Split(' ', 2, StringSplitOptions.RemoveEmptyEntries);
            return parts.Length > 0 ? parts[0].Trim('"') : null;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// 等待出现不属于已占用集合的新 WPS 进程。
    /// </summary>
    /// <param name="before">启动前的进程集合。</param>
    /// <param name="started">Process.Start 返回值。</param>
    /// <param name="owned">本软件已占用的进程。</param>
    /// <returns>新进程；超时则为 null。</returns>
    private static Process? WaitForNewProcess(HashSet<int> before, Process? started, IReadOnlyCollection<int> owned)
    {
        if (started is { HasExited: false } && !owned.Contains(started.Id))
        {
            return started;
        }

        var deadline = DateTime.UtcNow.AddSeconds(8);
        while (DateTime.UtcNow < deadline)
        {
            foreach (var name in new[] { "wps", "wpp", "wpsoffice" })
            {
                foreach (var process in Process.GetProcessesByName(name))
                {
                    if (!before.Contains(process.Id) && !owned.Contains(process.Id) && !process.HasExited)
                    {
                        return process;
                    }
                }
            }

            Thread.Sleep(120);
        }

        return started is { HasExited: false } ? started : null;
    }

    /// <summary>
    /// 创建 KWPP.Application。本机 WPS 12.1 的 NativeOM/ROT 绑不到新进程，必须走 CreateInstance。
    /// </summary>
    /// <returns>演示 Application。</returns>
    private static dynamic CreateKwppApplication()
    {
        var type = Type.GetTypeFromProgID("KWPP.Application", throwOnError: false);
        if (type is null)
        {
            throw new InvalidOperationException("未注册 KWPP.Application，请先安装 WPS 演示。");
        }

        var created = Activator.CreateInstance(type);
        if (created is null)
        {
            throw new InvalidOperationException("无法创建 WPS 演示实例。");
        }

        return created;
    }

    /// <summary>
    /// 等待真正的演示引擎 wpp.exe（wps.exe /wpp 只是启动器）。
    /// </summary>
    /// <param name="before">CreateInstance 之前的引擎进程。</param>
    /// <returns>引擎进程；找不到则为 null。</returns>
    private static Process? WaitForEngineProcess(HashSet<int> before)
    {
        var deadline = DateTime.UtcNow.AddSeconds(12);
        Process? fallback = null;
        while (DateTime.UtcNow < deadline)
        {
            foreach (var process in Process.GetProcessesByName("wpp"))
            {
                if (process.HasExited)
                {
                    continue;
                }

                fallback = process;
                if (!before.Contains(process.Id))
                {
                    return process;
                }
            }

            Thread.Sleep(150);
        }

        return fallback;
    }

    /// <summary>
    /// 当前本机演示引擎进程 ID。
    /// </summary>
    /// <returns>wpp.exe 进程 ID 集合。</returns>
    private static HashSet<int> SnapshotEngineProcessIds()
    {
        var ids = new HashSet<int>();
        foreach (var process in Process.GetProcessesByName("wpp"))
        {
            ids.Add(process.Id);
        }

        return ids;
    }

    /// <summary>
    /// 从该进程的窗口取出 Application，避免 GetActiveObject 拿到第一路实例。
    /// </summary>
    /// <param name="processId">目标进程 ID。</param>
    /// <returns>Application；失败则为 null。</returns>
    private static dynamic? TryBindFromProcess(int processId)
    {
        foreach (var hwnd in NativeWindowHelper.EnumerateProcessTopWindows(processId))
        {
            foreach (var node in NativeWindowHelper.CollectWindowTree(hwnd))
            {
                if (TryApplicationFromNativeOm(node) is { } application)
                {
                    return application;
                }
            }
        }

        return null;
    }

    /// <summary>
    /// 从 Running Object Table 取出属于该进程的 KWPP.Application。
    /// </summary>
    /// <param name="processId">目标进程 ID。</param>
    /// <returns>Application；失败则为 null。</returns>
    private static dynamic? TryBindFromRot(int processId)
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
                    if (obj is null || !LooksLikeKwpp(obj))
                    {
                        continue;
                    }

                    if (TryReadProcessId(obj) == processId)
                    {
                        return obj;
                    }
                }
                catch
                {
                    // 下一条 ROT 项
                }
                finally
                {
                    Marshal.ReleaseComObject(moniker);
                }
            }
        }
        catch
        {
            return null;
        }
        finally
        {
            Marshal.ReleaseComObject(rot);
        }

        return null;
    }

    /// <summary>
    /// 从窗口 NativeOM 取出 KWPP.Application。
    /// </summary>
    /// <param name="hwnd">窗口。</param>
    /// <returns>Application；失败则为 null。</returns>
    private static dynamic? TryApplicationFromNativeOm(IntPtr hwnd)
    {
        var raw = NativeWindowHelper.TryGetNativeOm(hwnd);
        if (raw is null)
        {
            return null;
        }

        dynamic om = raw;
        try
        {
            var app = om.Application;
            if (app is not null && LooksLikeKwpp(app))
            {
                return app;
            }
        }
        catch
        {
            // 不是 Window 对象
        }

        try
        {
            if (LooksLikeKwpp(om))
            {
                return om;
            }
        }
        catch
        {
            // 不是 Application
        }

        return null;
    }

    /// <summary>
    /// 判断对象是否像 WPS 演示 Application。
    /// </summary>
    /// <param name="candidate">COM 对象。</param>
    /// <returns>能访问 Presentations 则为 true。</returns>
    private static bool LooksLikeKwpp(object candidate)
    {
        try
        {
            dynamic app = candidate;
            _ = app.Presentations;
            return true;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// 关闭提示并尽量显示 Application。
    /// </summary>
    /// <param name="application">COM 对象。</param>
    private static void Configure(dynamic application)
    {
        try
        {
            application.Visible = true;
        }
        catch
        {
            // 个别版本 Visible 只读
        }

        try
        {
            application.DisplayAlerts = MsoFalse;
        }
        catch
        {
            // 忽略
        }
    }

    /// <summary>
    /// 从 Application 窗口推断进程 ID。
    /// </summary>
    /// <param name="application">COM 对象。</param>
    /// <returns>进程 ID；失败则为 0。</returns>
    private static int TryReadProcessId(object application)
    {
        dynamic app = application;
        try
        {
            foreach (dynamic window in app.Windows)
            {
                try
                {
                    var hwnd = new IntPtr(Convert.ToInt64(window.HWND));
                    var pid = NativeWindowHelper.GetWindowProcessId(hwnd);
                    if (pid > 0)
                    {
                        return pid;
                    }
                }
                catch
                {
                    // 下一扇窗
                }
            }
        }
        catch
        {
            // 无 Windows 集合
        }

        try
        {
            var hwnd = new IntPtr(Convert.ToInt64(app.HWND));
            return NativeWindowHelper.GetWindowProcessId(hwnd);
        }
        catch
        {
            return 0;
        }
    }

    /// <summary>
    /// 当前本机 WPS 相关进程 ID。
    /// </summary>
    /// <returns>进程 ID 集合。</returns>
    private static HashSet<int> SnapshotWpsProcessIds()
    {
        var ids = new HashSet<int>();
        foreach (var name in new[] { "wps", "wpp", "wpsoffice" })
        {
            foreach (var process in Process.GetProcessesByName(name))
            {
                ids.Add(process.Id);
            }
        }

        return ids;
    }

    [DllImport("ole32.dll")]
    private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable pprot);
}
