using System.Windows;
using System.Windows.Media.Animation;
using MultiPPTController.Services;
using MultiPPTController.ViewModels;
using Wpf.Ui.Controls;

namespace MultiPPTController;

/// <summary>
/// 主控制台：按设计稿还原的导入、分屏与放映界面。
/// </summary>
public partial class MainWindow : FluentWindow
{
    private readonly GlobalHotkeyService _hotkeys = new();
    private Storyboard? _breath;

    /// <summary>
    /// 创建窗口并挂接视图模型。
    /// </summary>
    public MainWindow()
    {
        InitializeComponent();
        var viewModel = new MainViewModel();
        DataContext = viewModel;
        Loaded += OnLoaded;
        Closed += OnClosed;
        viewModel.PropertyChanged += OnViewModelChanged;
    }

    /// <summary>
    /// 当前视图模型。
    /// </summary>
    private MainViewModel ViewModel => (MainViewModel)DataContext;

    /// <summary>
    /// 启动状态点动效与热键钩子。
    /// </summary>
    /// <param name="sender">窗口。</param>
    /// <param name="e">路由参数。</param>
    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        _breath = (Storyboard)Resources["BreathStoryboard"];
        if (SystemParameters.ClientAreaAnimation)
        {
            _breath.Begin();
        }

        _hotkeys.Attach(this);
        _hotkeys.NextRequested += () => ViewModel.NextCommand.Execute(null);
        _hotkeys.PreviousRequested += () => ViewModel.PreviousCommand.Execute(null);
        _hotkeys.StopRequested += () => ViewModel.StopCommand.Execute(null);
        ThemeGlyph.Text = ThemeService.IsDark ? "☽" : "☀";
    }

    /// <summary>
    /// 结束放映并释放热键。
    /// </summary>
    /// <param name="sender">窗口。</param>
    /// <param name="e">关闭参数。</param>
    private void OnClosed(object? sender, EventArgs e)
    {
        _hotkeys.Dispose();
        ViewModel.Dispose();
    }

    /// <summary>
    /// 打开多选文件框并导入。
    /// </summary>
    /// <param name="sender">按钮。</param>
    /// <param name="e">路由参数。</param>
    private void OnImportClick(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Multiselect = true,
            Filter = "演示文稿|*.pptx;*.ppt;*.pptm;*.dps;*.dpt;*.dptx;*.pdf|PDF|*.pdf|所有文件|*.*"
        };

        if (dialog.ShowDialog(this) == true)
        {
            ViewModel.ImportCommand.Execute(dialog.FileNames);
        }
    }

    /// <summary>
    /// 允许拖入演示文稿文件。
    /// </summary>
    /// <param name="sender">窗口。</param>
    /// <param name="e">拖放参数。</param>
    private void OnWindowDragOver(object sender, DragEventArgs e)
    {
        e.Effects = HasPresentationFiles(e) ? DragDropEffects.Copy : DragDropEffects.None;
        e.Handled = true;
    }

    /// <summary>
    /// 处理拖入的文件路径。
    /// </summary>
    /// <param name="sender">窗口。</param>
    /// <param name="e">拖放参数。</param>
    private void OnWindowDrop(object sender, DragEventArgs e)
    {
        if (e.Data.GetData(DataFormats.FileDrop) is string[] paths)
        {
            ViewModel.ImportCommand.Execute(paths);
        }
    }

    /// <summary>
    /// 在设计稿深色 / 浅色模式之间切换。
    /// </summary>
    /// <param name="sender">按钮。</param>
    /// <param name="e">路由参数。</param>
    private void OnToggleThemeClick(object sender, RoutedEventArgs e)
    {
        ThemeService.Apply(!ThemeService.IsDark);
        ThemeGlyph.Text = ThemeService.IsDark ? "☽" : "☀";
    }

    /// <summary>
    /// 打开关于窗口。
    /// </summary>
    /// <param name="sender">按钮。</param>
    /// <param name="e">路由参数。</param>
    private void OnAboutClick(object sender, RoutedEventArgs e)
    {
        var about = new AboutWindow { Owner = this };
        about.ShowDialog();
    }

    /// <summary>
    /// 放映状态变化时开关全局热键。
    /// </summary>
    /// <param name="sender">视图模型。</param>
    /// <param name="e">属性参数。</param>
    private void OnViewModelChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs e)
    {
        if (e.PropertyName != nameof(MainViewModel.IsPlaying))
        {
            return;
        }

        if (ViewModel.IsPlaying)
        {
            _hotkeys.StartCapture();
            HideConsole();
        }
        else
        {
            _hotkeys.StopCapture();
            RestoreConsole();
        }
    }

    /// <summary>
    /// 拖放数据是否包含支持的演示文件。
    /// </summary>
    /// <param name="e">拖放参数。</param>
    /// <returns>包含则为 true。</returns>
    private static bool HasPresentationFiles(DragEventArgs e)
    {
        if (!e.Data.GetDataPresent(DataFormats.FileDrop) ||
            e.Data.GetData(DataFormats.FileDrop) is not string[] paths)
        {
            return false;
        }

        var allowed = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            ".pptx", ".ppt", ".pptm", ".dps", ".dpt", ".dptx", ".pdf"
        };
        return paths.Any(path => allowed.Contains(Path.GetExtension(path)));
    }

    /// <summary>
    /// 开播后收起控制台，翻页用主屏播放条或方向键，Esc 结束。
    /// </summary>
    private void HideConsole()
    {
        Hide();
        ShowInTaskbar = false;
    }

    /// <summary>
    /// 结束放映后恢复主控制台。
    /// </summary>
    private void RestoreConsole()
    {
        ShowInTaskbar = true;
        Show();
        WindowState = WindowState.Normal;
        Activate();
    }
}
