using System.Globalization;
using System.Windows.Data;

namespace MultiPPTController.Converters;

/// <summary>
/// 将 bool 取反，用于放映中锁定片库。
/// </summary>
public sealed class InverseBoolConverter : IValueConverter
{
    /// <summary>
    /// 取反布尔值。
    /// </summary>
    /// <param name="value">源值。</param>
    /// <param name="targetType">目标类型。</param>
    /// <param name="parameter">未使用。</param>
    /// <param name="culture">区域。</param>
    /// <returns>取反后的 bool；无法转换则为 true。</returns>
    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        return value is bool flag ? !flag : true;
    }

    /// <summary>
    /// 双向绑定时再次取反。
    /// </summary>
    /// <param name="value">目标值。</param>
    /// <param name="targetType">源类型。</param>
    /// <param name="parameter">未使用。</param>
    /// <param name="culture">区域。</param>
    /// <returns>取反后的 bool。</returns>
    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        return value is bool flag ? !flag : false;
    }
}
