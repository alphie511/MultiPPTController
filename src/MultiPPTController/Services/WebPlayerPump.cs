using System.Windows.Threading;

namespace MultiPPTController.Services;

/// <summary>
/// 独立 STA 消息泵，专给 WebView2 主屏窗使用，避免冒烟时卡死 UI 线程。
/// </summary>
public sealed class WebPlayerPump : IDisposable
{
    private readonly Thread _thread;
    private Dispatcher? _dispatcher;
    private bool _disposed;

    /// <summary>
    /// 启动后台 STA 线程并进入 Dispatcher 循环。
    /// </summary>
    public WebPlayerPump()
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
            Name = "WebPlayerSta"
        };
        _thread.SetApartmentState(ApartmentState.STA);
        _thread.Start();
        ready.Wait();
    }

    /// <summary>
    /// 在 Web 线程同步执行并返回结果。
    /// </summary>
    /// <typeparam name="T">返回值类型。</typeparam>
    /// <param name="func">要执行的函数。</param>
    /// <returns>函数结果。</returns>
    public T Invoke<T>(Func<T> func)
    {
        ArgumentNullException.ThrowIfNull(func);
        EnsureAlive();
        return _dispatcher!.CheckAccess() ? func() : _dispatcher.Invoke(func);
    }

    /// <summary>
    /// 在 Web 线程同步执行无返回值操作。
    /// </summary>
    /// <param name="action">要执行的操作。</param>
    public void Invoke(Action action)
    {
        ArgumentNullException.ThrowIfNull(action);
        EnsureAlive();
        if (_dispatcher!.CheckAccess())
        {
            action();
            return;
        }

        _dispatcher.Invoke(action);
    }

    /// <summary>
    /// 在 Web 线程排队执行异步操作，并等待完成。
    /// </summary>
    /// <param name="action">异步操作。</param>
    /// <returns>完成任务。</returns>
    public Task InvokeAsync(Func<Task> action)
    {
        ArgumentNullException.ThrowIfNull(action);
        EnsureAlive();
        var completion = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        _ = _dispatcher!.BeginInvoke(async () =>
        {
            try
            {
                await action().ConfigureAwait(true);
                completion.TrySetResult();
            }
            catch (Exception ex)
            {
                completion.TrySetException(ex);
            }
        });
        return completion.Task;
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
