using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;

namespace MultiPPTController.Services;

/// <summary>
/// 放映期间注册全局热键，避免焦点落在 WPS 窗口后控制台收不到按键。
/// </summary>
public sealed class GlobalHotkeyService : IDisposable
{
    private const int WmHotkey = 0x0312;
    private const uint ModNoRepeat = 0x4000;
    private const int VkLeft = 0x25;
    private const int VkUp = 0x26;
    private const int VkRight = 0x27;
    private const int VkDown = 0x28;
    private const int VkSpace = 0x20;
    private const int VkPrior = 0x21;
    private const int VkNext = 0x22;
    private const int VkEscape = 0x1B;

    private HwndSource? _source;
    private IntPtr _hwnd;
    private bool _registered;
    private bool _disposed;

    /// <summary>下一页 / 下一步动画。</summary>
    public event Action? NextRequested;

    /// <summary>上一页 / 上一步动画。</summary>
    public event Action? PreviousRequested;

    /// <summary>结束放映。</summary>
    public event Action? StopRequested;

    /// <summary>
    /// 绑定到控制台窗口的消息循环。
    /// </summary>
    /// <param name="window">主窗口。</param>
    public void Attach(Window window)
    {
        var helper = new WindowInteropHelper(window);
        helper.EnsureHandle();
        _hwnd = helper.Handle;
        _source = HwndSource.FromHwnd(_hwnd);
        _source?.AddHook(WndProc);
    }

    /// <summary>
    /// 开始拦截翻页与结束热键。
    /// </summary>
    public void StartCapture()
    {
        if (_registered || _hwnd == IntPtr.Zero)
        {
            return;
        }

        Register(1, VkRight);
        Register(2, VkSpace);
        Register(3, VkNext);
        Register(4, VkDown);
        Register(5, VkLeft);
        Register(6, VkPrior);
        Register(7, VkUp);
        Register(8, VkEscape);
        _registered = true;
    }

    /// <summary>
    /// 停止拦截，把按键还给系统。
    /// </summary>
    public void StopCapture()
    {
        if (!_registered || _hwnd == IntPtr.Zero)
        {
            return;
        }

        for (var id = 1; id <= 8; id++)
        {
            UnregisterHotKey(_hwnd, id);
        }

        _registered = false;
    }

    /// <summary>
    /// 注销热键并移除钩子。
    /// </summary>
    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;
        StopCapture();
        _source?.RemoveHook(WndProc);
        _source = null;
    }

    /// <summary>
    /// 注册单个热键。
    /// </summary>
    /// <param name="id">热键编号。</param>
    /// <param name="vk">虚拟键码。</param>
    private void Register(int id, int vk)
    {
        RegisterHotKey(_hwnd, id, ModNoRepeat, (uint)vk);
    }

    /// <summary>
    /// 分发 WM_HOTKEY。
    /// </summary>
    private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
    {
        if (msg != WmHotkey)
        {
            return IntPtr.Zero;
        }

        switch (wParam.ToInt32())
        {
            case 1:
            case 2:
            case 3:
            case 4:
                NextRequested?.Invoke();
                handled = true;
                break;
            case 5:
            case 6:
            case 7:
                PreviousRequested?.Invoke();
                handled = true;
                break;
            case 8:
                StopRequested?.Invoke();
                handled = true;
                break;
        }

        return IntPtr.Zero;
    }

    [DllImport("user32.dll")]
    private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

    [DllImport("user32.dll")]
    private static extern bool UnregisterHotKey(IntPtr hWnd, int id);
}
