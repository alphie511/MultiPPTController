using System.Windows;
using MultiPPTController.Services;

namespace MultiPPTController;

/// <summary>
/// 应用程序入口，默认套用设计稿深色主题。
/// </summary>
public partial class App : System.Windows.Application
{
    /// <summary>
    /// 启动后立刻对齐设计稿深色模式。
    /// </summary>
    /// <param name="e">启动参数。</param>
    protected override void OnStartup(StartupEventArgs e)
    {
        if (e.Args.Any(item => string.Equals(item, "--smoke", StringComparison.OrdinalIgnoreCase)))
        {
            Environment.ExitCode = PlaybackSmoke.Run();
            Shutdown();
            return;
        }

        base.OnStartup(e);
        ThemeService.Apply(true);
    }
}
