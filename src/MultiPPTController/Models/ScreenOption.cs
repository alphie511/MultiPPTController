using CommunityToolkit.Mvvm.ComponentModel;

namespace MultiPPTController.Models;

/// <summary>
/// 片库下拉框中的一块屏幕选项，可用性相对「当前这份 PPT」计算。
/// </summary>
public sealed partial class ScreenOption : ObservableObject
{
    /// <summary>
    /// 用物理屏幕创建下拉选项。
    /// </summary>
    /// <param name="screen">已连接显示器。</param>
    public ScreenOption(DisplayScreen screen)
    {
        Screen = screen;
        DeviceName = screen.DeviceName;
        Label = screen.Label;
    }

    /// <summary>对应显示器。</summary>
    public DisplayScreen Screen { get; }

    /// <summary>系统设备名。</summary>
    public string DeviceName { get; }

    /// <summary>界面展示名，例如「屏幕 1 · 主屏」。</summary>
    public string Label { get; }

    /// <summary>当前这份 PPT 是否可选该屏；被其他 PPT 占用则为 false。</summary>
    [ObservableProperty]
    private bool _isAvailable = true;

    /// <summary>当前这份 PPT 是否已勾选该屏。</summary>
    [ObservableProperty]
    private bool _isSelected;

    /// <summary>勾选变化时通知主视图（同步期间可置空以免回写）。</summary>
    public Action<ScreenOption, bool>? SelectionChanged { get; set; }

    /// <summary>
    /// 勾选变化时回写片库绑定。
    /// </summary>
    /// <param name="value">是否选中。</param>
    partial void OnIsSelectedChanged(bool value)
    {
        SelectionChanged?.Invoke(this, value);
    }

    /// <summary>
    /// 静默设置勾选，不触发回写（用于按已绑定列表刷新 UI）。
    /// </summary>
    /// <param name="value">是否选中。</param>
    public void SetSelectedSilent(bool value)
    {
        var callback = SelectionChanged;
        SelectionChanged = null;
        IsSelected = value;
        SelectionChanged = callback;
    }
}
