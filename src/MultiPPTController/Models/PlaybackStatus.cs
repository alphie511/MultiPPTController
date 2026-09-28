namespace MultiPPTController.Models;

/// <summary>
/// 编排器对外暴露的放映阶段。
/// </summary>
public enum PlaybackStatus
{
    /// <summary>空闲，可导入与编排。</summary>
    Idle,

    /// <summary>正在打开 WPS 并进入放映。</summary>
    Starting,

    /// <summary>各屏正在放映。</summary>
    Playing,

    /// <summary>正在退出放映并关闭文稿。</summary>
    Stopping
}
