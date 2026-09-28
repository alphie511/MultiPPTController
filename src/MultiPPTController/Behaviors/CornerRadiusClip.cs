using System.Windows;
using System.Windows.Media;

namespace MultiPPTController.Behaviors;

/// <summary>
/// 按 Border 的 CornerRadius 裁剪子内容，避免封面图把圆角盖成直角。
/// </summary>
public static class CornerRadiusClip
{
    /// <summary>是否启用圆角裁剪。</summary>
    public static readonly DependencyProperty EnabledProperty = DependencyProperty.RegisterAttached(
        "Enabled",
        typeof(bool),
        typeof(CornerRadiusClip),
        new PropertyMetadata(false, OnEnabledChanged));

    /// <summary>
    /// 读取是否启用。
    /// </summary>
    /// <param name="element">目标元素。</param>
    /// <returns>已启用则为 true。</returns>
    public static bool GetEnabled(DependencyObject element) => (bool)element.GetValue(EnabledProperty);

    /// <summary>
    /// 设置是否启用。
    /// </summary>
    /// <param name="element">目标元素。</param>
    /// <param name="value">是否启用。</param>
    public static void SetEnabled(DependencyObject element, bool value) => element.SetValue(EnabledProperty, value);

    /// <summary>
    /// 订阅尺寸变化并更新裁剪几何。
    /// </summary>
    /// <param name="sender">附加对象。</param>
    /// <param name="e">属性参数。</param>
    private static void OnEnabledChanged(DependencyObject sender, DependencyPropertyChangedEventArgs e)
    {
        if (sender is not FrameworkElement element)
        {
            return;
        }

        element.SizeChanged -= OnSizeChanged;
        if ((bool)e.NewValue)
        {
            element.SizeChanged += OnSizeChanged;
            ApplyClip(element);
        }
        else
        {
            element.Clip = null;
        }
    }

    /// <summary>
    /// 尺寸变化后重算圆角裁剪。
    /// </summary>
    /// <param name="sender">元素。</param>
    /// <param name="e">尺寸参数。</param>
    private static void OnSizeChanged(object sender, SizeChangedEventArgs e)
    {
        if (sender is FrameworkElement element)
        {
            ApplyClip(element);
        }
    }

    /// <summary>
    /// 用当前宽高生成圆角矩形裁剪。
    /// </summary>
    /// <param name="element">目标元素。</param>
    private static void ApplyClip(FrameworkElement element)
    {
        if (element.ActualWidth <= 0 || element.ActualHeight <= 0)
        {
            return;
        }

        var radius = 15.0;
        if (element is System.Windows.Controls.Border border)
        {
            radius = Math.Max(border.CornerRadius.TopLeft, 0);
        }

        element.Clip = new RectangleGeometry(
            new Rect(0, 0, element.ActualWidth, element.ActualHeight),
            radius,
            radius);
    }
}
