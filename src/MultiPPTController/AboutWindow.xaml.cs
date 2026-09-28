using System.Windows;
using Wpf.Ui.Controls;

namespace MultiPPTController;

/// <summary>
/// 关于窗口，版式对齐设计稿。
/// </summary>
public partial class AboutWindow : FluentWindow
{
    /// <summary>
    /// 创建关于窗口。
    /// </summary>
    public AboutWindow()
    {
        InitializeComponent();
    }

    /// <summary>
    /// 提示当前已是最新版本。
    /// </summary>
    /// <param name="sender">按钮。</param>
    /// <param name="e">路由参数。</param>
    private void OnCheckUpdateClick(object sender, RoutedEventArgs e)
    {
        System.Windows.MessageBox.Show(this, "当前已是最新版本 v1.0.0。", "检查更新",
            System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
    }

    /// <summary>
    /// 关闭关于窗口。
    /// </summary>
    /// <param name="sender">按钮。</param>
    /// <param name="e">路由参数。</param>
    private void OnCloseClick(object sender, RoutedEventArgs e)
    {
        Close();
    }
}
