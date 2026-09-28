using System.Windows.Media;
using CommunityToolkit.Mvvm.ComponentModel;
using MultiPPTController.Models;

namespace MultiPPTController.ViewModels;

/// <summary>
/// 右侧屏幕编排卡片：屏幕信息 + 已绑定 PPT 封面。
/// </summary>
public sealed partial class ScreenCardViewModel : ObservableObject
{
    /// <summary>
    /// 绑定一块物理屏幕。
    /// </summary>
    /// <param name="screen">显示器。</param>
    public ScreenCardViewModel(DisplayScreen screen)
    {
        Screen = screen;
    }

    /// <summary>对应显示器。</summary>
    public DisplayScreen Screen { get; }

    /// <summary>屏幕展示名。</summary>
    public string Label => Screen.Label;

    /// <summary>分辨率文案。</summary>
    public string ResolutionText => Screen.ResolutionText;

    /// <summary>已绑定文稿封面；未绑定则为 null。</summary>
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
}
