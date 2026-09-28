using System.Windows;
using Wpf.Ui.Appearance;

namespace MultiPPTController.Services;

/// <summary>
/// 在设计稿的深色 / 浅色两套资源之间切换。
/// </summary>
public static class ThemeService
{
    /// <summary>当前是否为深色模式。</summary>
    public static bool IsDark { get; private set; } = true;

    /// <summary>
    /// 应用指定主题，并替换 WPF-UI 与本地色板。
    /// </summary>
    /// <param name="dark">true 为深色模式。</param>
    public static void Apply(bool dark)
    {
        IsDark = dark;
        var app = Application.Current;
        var merged = app.Resources.MergedDictionaries;
        for (var i = merged.Count - 1; i >= 0; i--)
        {
            var source = merged[i].Source?.OriginalString ?? string.Empty;
            if (source.EndsWith("Dark.xaml", StringComparison.OrdinalIgnoreCase) ||
                source.EndsWith("Light.xaml", StringComparison.OrdinalIgnoreCase))
            {
                merged.RemoveAt(i);
            }
        }

        var file = dark ? "Themes/Dark.xaml" : "Themes/Light.xaml";
        merged.Add(new ResourceDictionary
        {
            Source = new Uri(file, UriKind.Relative)
        });

        ApplicationThemeManager.Apply(dark ? ApplicationTheme.Dark : ApplicationTheme.Light);
    }
}
