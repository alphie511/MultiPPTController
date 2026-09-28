namespace MultiPPTController.Models;

/// <summary>
/// 表示一块已连接的显示器，用于绑定 PPT 输出目标。
/// </summary>
/// <param name="DeviceName">系统设备名，例如 \\.\DISPLAY1。</param>
/// <param name="Label">界面展示名，例如「屏幕 1 · 主屏」。</param>
/// <param name="Bounds">虚拟桌面坐标下的像素矩形。</param>
/// <param name="IsPrimary">是否为主显示器。</param>
public sealed record DisplayScreen(
    string DeviceName,
    string Label,
    System.Drawing.Rectangle Bounds,
    bool IsPrimary)
{
    /// <summary>
    /// 分辨率文案，供卡片副标题使用。
    /// </summary>
    public string ResolutionText => $"{Bounds.Width} × {Bounds.Height}";
}
