using System.Windows.Media;

namespace HarmonicSheet.Services;

public static class ColorTheme
{
    public static string CurrentTheme { get; set; } = "ModernMinimal";

    public static ThemeColors GetColors()
    {
        return CurrentTheme switch
        {
            "GoogleSheets" => new ThemeColors
            {
                // すべてグレー統一
                ToolbarButtonBackground = "#F3F4F6",
                ToolbarButtonForeground = "#1F2937",
                CalcButtonBackground = "#F3F4F6",
                CalcButtonForeground = "#1F2937",
                CalcLabelBackground = "#E5E7EB",
                CalcLabelForeground = "#1F2937",
                CalcRowBackground = "#FFFFFF",
                CalcRowBorder = "#E5E7EB",
                PrimaryButtonBackground = "#3B82F6", // 保存とAI実行のみ青
                PrimaryButtonForeground = "#FFFFFF"
            },
            "SeniorFriendly" => new ThemeColors
            {
                // ベージュ・優しい色
                ToolbarButtonBackground = "#FFFFFF",
                ToolbarButtonForeground = "#44403C",
                CalcButtonBackground = "#FFFFFF",
                CalcButtonForeground = "#44403C",
                CalcLabelBackground = "#F5F5F4",
                CalcLabelForeground = "#44403C",
                CalcRowBackground = "#FAFAF9",
                CalcRowBorder = "#E7E5E4",
                PrimaryButtonBackground = "#3B82F6",
                PrimaryButtonForeground = "#FFFFFF"
            },
            "ModernMinimal" or _ => new ThemeColors
            {
                // 薄い青アクセント
                ToolbarButtonBackground = "#F3F4F6",
                ToolbarButtonForeground = "#1F2937",
                CalcButtonBackground = "#EFF6FF", // 薄い青
                CalcButtonForeground = "#1E40AF",
                CalcLabelBackground = "#DBEAFE",
                CalcLabelForeground = "#1E40AF",
                CalcRowBackground = "#F9FAFB",
                CalcRowBorder = "#BFDBFE",
                PrimaryButtonBackground = "#3B82F6",
                PrimaryButtonForeground = "#FFFFFF"
            }
        };
    }
}

public class ThemeColors
{
    public string ToolbarButtonBackground { get; set; } = "#F3F4F6";
    public string ToolbarButtonForeground { get; set; } = "#1F2937";
    public string CalcButtonBackground { get; set; } = "#EFF6FF";
    public string CalcButtonForeground { get; set; } = "#1E40AF";
    public string CalcLabelBackground { get; set; } = "#DBEAFE";
    public string CalcLabelForeground { get; set; } = "#1E40AF";
    public string CalcRowBackground { get; set; } = "#F9FAFB";
    public string CalcRowBorder { get; set; } = "#BFDBFE";
    public string PrimaryButtonBackground { get; set; } = "#3B82F6";
    public string PrimaryButtonForeground { get; set; } = "#FFFFFF";

    public Brush ToolbarButtonBrush => new SolidColorBrush((Color)ColorConverter.ConvertFromString(ToolbarButtonBackground));
    public Brush ToolbarForegroundBrush => new SolidColorBrush((Color)ColorConverter.ConvertFromString(ToolbarButtonForeground));
    public Brush CalcButtonBrush => new SolidColorBrush((Color)ColorConverter.ConvertFromString(CalcButtonBackground));
    public Brush CalcForegroundBrush => new SolidColorBrush((Color)ColorConverter.ConvertFromString(CalcButtonForeground));
    public Brush CalcLabelBrush => new SolidColorBrush((Color)ColorConverter.ConvertFromString(CalcLabelBackground));
    public Brush CalcLabelForegroundBrush => new SolidColorBrush((Color)ColorConverter.ConvertFromString(CalcLabelForeground));
    public Brush CalcRowBrush => new SolidColorBrush((Color)ColorConverter.ConvertFromString(CalcRowBackground));
    public Brush CalcRowBorderBrush => new SolidColorBrush((Color)ColorConverter.ConvertFromString(CalcRowBorder));
    public Brush PrimaryBrush => new SolidColorBrush((Color)ColorConverter.ConvertFromString(PrimaryButtonBackground));
    public Brush PrimaryForegroundBrush => new SolidColorBrush((Color)ColorConverter.ConvertFromString(PrimaryButtonForeground));
}
