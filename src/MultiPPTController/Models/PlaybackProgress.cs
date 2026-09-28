namespace MultiPPTController.Models;

/// <summary>
/// 一路放映的进度：页码 + 当前页已发生的单击动画次数。
/// </summary>
/// <param name="SlideIndex">当前幻灯片序号（从 1 起）；0 表示未能读取。</param>
/// <param name="ClickIndex">当前页单击动画进度；读不到则为 0。</param>
public readonly record struct PlaybackProgress(int SlideIndex, int ClickIndex) : IComparable<PlaybackProgress>
{
    /// <summary>
    /// 先比页码，再比单击进度。
    /// </summary>
    /// <param name="other">另一路进度。</param>
    /// <returns>比较结果。</returns>
    public int CompareTo(PlaybackProgress other)
    {
        var slide = SlideIndex.CompareTo(other.SlideIndex);
        return slide != 0 ? slide : ClickIndex.CompareTo(other.ClickIndex);
    }

    /// <summary>是否读到了有效页码。</summary>
    public bool HasSlide => SlideIndex > 0;

    /// <inheritdoc />
    public override string ToString() => HasSlide ? $"第 {SlideIndex} 页 · 单击 {ClickIndex}" : "未知";
}
