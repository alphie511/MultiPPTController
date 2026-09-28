using System.Windows.Threading;

namespace MultiPPTController.Services;

/// <summary>
/// 独立 STA 消息泵，保证所有 WPS COM 调用发生在同一线程。
/// </summary>
public sealed class StaComPump : IDisposable
{
    private readonly Thread _thread;
    private Dispatcher? _dispatcher;
    private bool _disposed;

    /// <summary>
    /// 启动后台 STA 线程并进入 Dispatcher 循环。
    /// </summary>
    public StaComPump()
    {
        using var ready = new ManualResetEventSlim(false);
        _thread = new Thread(() =>
        {
            _dispatcher = Dispatcher.CurrentDispatcher;
            ready.Set();
            Dispatcher.Run();
        })
        {
            IsBackground = true,
            Name = "WpsComSta"
        };
        _thread.SetApartmentState(ApartmentState.STA);
        _thread.Start();
        ready.Wait();
    }

    /// <summary>
    /// 在 COM 线程执行无返回值操作。
    /// </summary>
    /// <param name="action">要执行的操作。</param>
    /// <returns>完成任务。</returns>
    public Task InvokeAsync(Action action)
    {
        ArgumentNullException.ThrowIfNull(action);
        EnsureAlive();
        return _dispatcher!.InvokeAsync(action).Task;
    }

    /// <summary>
    /// 在 COM 线程执行并返回结果。
    /// </summary>
    /// <typeparam name="T">返回值类型。</typeparam>
    /// <param name="func">要执行的函数。</param>
    /// <returns>函数结果任务。</returns>
    public Task<T> InvokeAsync<T>(Func<T> func)
    {
        ArgumentNullException.ThrowIfNull(func);
        EnsureAlive();
        return _dispatcher!.InvokeAsync(func).Task;
    }

    /// <summary>
    /// 关闭消息泵并等待线程退出。
    /// </summary>
    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;
        _dispatcher?.InvokeShutdown();
        if (_thread.IsAlive)
        {
            _thread.Join(TimeSpan.FromSeconds(3));
        }
    }

    /// <summary>
    /// 确认泵尚未释放。
    /// </summary>
    private void EnsureAlive()
    {
        ObjectDisposedException.ThrowIf(_disposed || _dispatcher is null, this);
    }
}
