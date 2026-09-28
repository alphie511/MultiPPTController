using MultiPPTController.Models;
using Screen = System.Windows.Forms.Screen;

namespace MultiPPTController.Services;

/// <summary>
/// 枚举当前 Windows 扩展桌面上的显示器。
/// </summary>
public sealed class MonitorService
{
    /// <summary>
    /// 读取当前已连接屏幕列表，按虚拟桌面 X 坐标排序。
    /// </summary>
    /// <returns>屏幕快照集合。</returns>
    public IReadOnlyList<DisplayScreen> GetScreens()
    {
        return Screen.AllScreens
            .OrderBy(item => item.Bounds.X)
            .ThenBy(item => item.Bounds.Y)
            .Select((item, index) => new DisplayScreen(
                item.DeviceName,
                item.Primary ? $"屏幕 {index + 1} · 主屏" : $"屏幕 {index + 1}",
                NativeWindowHelper.GetPhysicalMonitorBounds(item.Bounds),
                item.Primary))
            .ToList();
    }
}
