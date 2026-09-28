using MultiPPTController.Models;

namespace MultiPPTController.Services;

/// <summary>
/// 单块屏幕上的一路放映：主屏 Web、副屏 WPS 共用同一套步进接口。
/// </summary>
public interface IScreenPlayer
{
    /// <summary>对应片库条目。</summary>
    DeckItem Deck { get; }

    /// <summary>输出屏幕。</summary>
    DisplayScreen Screen { get; }

    /// <summary>已捕获的放映窗口句柄；未捕获则为 Zero。</summary>
    IntPtr ShowHwnd { get; }

    /// <summary>
    /// 诊断：本路句柄与进度说明。
    /// </summary>
    /// <returns>单行状态。</returns>
    string DescribeShow();

    /// <summary>
    /// 读取当前页码（单击动画进度仅 WPS 可能有值）。
    /// </summary>
    /// <returns>进度快照；读失败时页码为 0。</returns>
    PlaybackProgress ReadProgress();

    /// <summary>
    /// 前进一步（页或单击，由具体引擎决定）。
    /// </summary>
    void Next();

    /// <summary>
    /// 后退一步。
    /// </summary>
    void Previous();

    /// <summary>
    /// 按目标屏物理像素重新铺满。
    /// </summary>
    /// <param name="occupiedByOthers">其它路已占用的窗口。</param>
    /// <param name="raiseZOrder">开播首次铺窗时置顶。</param>
    void RelayoutToScreen(IEnumerable<IntPtr>? occupiedByOthers = null, bool raiseZOrder = false);

    /// <summary>
    /// 开播后预热，避免第一下步进被当成激活吃掉。
    /// </summary>
    void WarmUp();

    /// <summary>
    /// 结束本路放映。
    /// </summary>
    void Stop();
}
