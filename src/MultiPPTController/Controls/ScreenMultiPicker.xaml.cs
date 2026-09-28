using System.Windows.Controls;

namespace MultiPPTController.Controls;

/// <summary>
/// 片库里的多选屏幕下拉：一份 PPT 可勾多块屏，被其他 PPT 占用的项灰显。
/// </summary>
public partial class ScreenMultiPicker : UserControl
{
    /// <summary>
    /// 初始化多选下拉外观。
    /// </summary>
    public ScreenMultiPicker()
    {
        InitializeComponent();
    }
}
