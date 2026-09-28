using System.Collections.ObjectModel;
using System.Windows.Media;
using CommunityToolkit.Mvvm.ComponentModel;

namespace MultiPPTController.Models;

/// <summary>
/// 一份已导入的演示文稿及其屏幕绑定（一份可绑多屏）。
/// </summary>
public sealed partial class DeckItem : ObservableObject
{
    /// <summary>
    /// 创建片库条目。
    /// </summary>
    /// <param name="filePath">本地绝对路径。</param>
    public DeckItem(string filePath)
    {
        FilePath = filePath;
        DisplayName = Path.GetFileNameWithoutExtension(filePath);
    }

    /// <summary>本地绝对路径。</summary>
    public string FilePath { get; }

    /// <summary>不含扩展名的文件名。</summary>
    public string DisplayName { get; }

    /// <summary>该份文稿下拉框里的屏幕选项（含占用禁用状态）。</summary>
    public ObservableCollection<ScreenOption> ScreenOptions { get; } = [];

    /// <summary>已绑定的显示器设备名；可多选。</summary>
    public ObservableCollection<string> AssignedDeviceNames { get; } = [];

    /// <summary>闭合框展示文案，例如「屏幕 1 · 主屏、屏幕 2」。</summary>
    [ObservableProperty]
    private string _assignmentSummary = "选择屏幕";

    /// <summary>幻灯片页数；打开失败时为 null。</summary>
    [ObservableProperty]
    private int? _slideCount;

    /// <summary>第一页封面；提取失败则为 null。</summary>
    [ObservableProperty]
    private ImageSource? _cover;

    /// <summary>是否已有封面可显示。</summary>
    [ObservableProperty]
    private bool _hasCover;

    /// <summary>
    /// 封面变化时同步 HasCover。
    /// </summary>
    /// <param name="value">新封面。</param>
    partial void OnCoverChanged(ImageSource? value)
    {
        HasCover = value is not null;
    }

    /// <summary>最近一次探测或放映的状态说明。</summary>
    [ObservableProperty]
    private string _status = "待分配屏幕";

    /// <summary>是否至少绑定了一块屏幕。</summary>
    public bool HasAssignment => AssignedDeviceNames.Count > 0;

    /// <summary>
    /// 是否已绑定指定设备。
    /// </summary>
    /// <param name="deviceName">系统设备名。</param>
    /// <returns>已绑定则为 true。</returns>
    public bool IsAssignedTo(string? deviceName)
    {
        return !string.IsNullOrWhiteSpace(deviceName) &&
               AssignedDeviceNames.Any(name =>
                   string.Equals(name, deviceName, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// 按当前屏幕列表刷新闭合框与状态文案。
    /// </summary>
    /// <param name="screens">已连接显示器。</param>
    public void RefreshAssignmentText(IEnumerable<DisplayScreen> screens)
    {
        var labels = AssignedDeviceNames
            .Select(name => screens.FirstOrDefault(screen => screen.DeviceName == name)?.Label)
            .Where(label => !string.IsNullOrWhiteSpace(label))
            .Cast<string>()
            .ToList();

        if (labels.Count == 0)
        {
            AssignmentSummary = "选择屏幕";
            Status = "待分配屏幕";
            return;
        }

        AssignmentSummary = string.Join("、", labels);
        Status = labels.Count == 1 ? $"将输出到 {labels[0]}" : $"将输出到 {labels.Count} 块屏幕";
    }
}
