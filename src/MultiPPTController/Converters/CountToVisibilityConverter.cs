using System.Collections;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace MultiPPTController.Converters;

/// <summary>
/// 将布尔值转为可见性；Invert 时取反。
/// </summary>
public sealed class CountToVisibilityConverter : IValueConverter
{
    /// <summary>为 true 时：false 才显示。</summary>
    public bool Invert { get; set; }

    /// <summary>
    /// 将布尔或数量转为可见性。
    /// </summary>
    /// <param name="value">bool、int 或集合。</param>
    /// <param name="targetType">目标类型。</param>
    /// <param name="parameter">未使用。</param>
    /// <param name="culture">区域。</param>
    /// <returns>Visible 或 Collapsed。</returns>
    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        var show = value switch
        {
            bool flag => flag,
            int number => number > 0,
            ICollection collection => collection.Count > 0,
            _ => false
        };
        if (Invert)
        {
            show = !show;
        }

        return show ? Visibility.Visible : Visibility.Collapsed;
    }

    /// <summary>
    /// 不支持反向转换。
    /// </summary>
    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        throw new NotSupportedException();
    }
}
