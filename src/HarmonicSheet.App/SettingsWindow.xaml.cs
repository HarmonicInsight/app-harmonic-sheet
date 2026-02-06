using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using HarmonicSheet.Services;
using HarmonicSheet.Views;
using Microsoft.Extensions.DependencyInjection;

namespace HarmonicSheet;

public partial class SettingsWindow : Window
{
    private static readonly string SettingsFilePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "HarmonicSheet",
        "settings.json");

    private readonly IAccessibilityService? _accessibilityService;

    public SettingsWindow()
    {
        InitializeComponent();

        // アクセシビリティサービスを取得
        _accessibilityService = App.Services?.GetService<IAccessibilityService>();

        LoadSettings();
    }

    private void LoadSettings()
    {
        try
        {
            // アクセシビリティ設定を読み込み
            if (_accessibilityService != null)
            {
                FontScaleSlider.Value = _accessibilityService.FontScale;
                SpeechRateSlider.Value = _accessibilityService.SpeechRate;
                HighContrastCheck.IsChecked = _accessibilityService.HighContrastMode;
                UpdateFontPreview();
            }

            // アプリ設定を読み込み
            if (File.Exists(SettingsFilePath))
            {
                var json = File.ReadAllText(SettingsFilePath);
                var settings = JsonSerializer.Deserialize<AppSettings>(json);

                if (settings != null)
                {
                    ApiKeyInput.Password = settings.ClaudeApiKey ?? string.Empty;
                    EmailInput.Text = settings.EmailAddress ?? string.Empty;
                    DisplayNameInput.Text = settings.DisplayName ?? string.Empty;
                    SmtpServerInput.Text = settings.SmtpServer ?? string.Empty;
                    SmtpPortInput.Text = settings.SmtpPort.ToString();
                    ImapServerInput.Text = settings.ImapServer ?? string.Empty;
                    ImapPortInput.Text = settings.ImapPort.ToString();

                    // カラーテーマを設定
                    var theme = settings.ColorTheme ?? "ModernMinimal";
                    foreach (ComboBoxItem item in ThemeComboBox.Items)
                    {
                        if (item.Tag?.ToString() == theme)
                        {
                            ThemeComboBox.SelectedItem = item;
                            break;
                        }
                    }

                    if (!string.IsNullOrEmpty(settings.ClaudeApiKey))
                    {
                        ApiKeyStatus.Text = "APIキーが設定されています";
                        ApiKeyStatus.Foreground = System.Windows.Media.Brushes.Green;
                    }
                }
            }
        }
        catch
        {
            // 設定ファイルの読み込みに失敗しても続行
        }
    }

    private void SaveSettings()
    {
        try
        {
            // アクセシビリティ設定を保存
            if (_accessibilityService != null)
            {
                _accessibilityService.FontScale = FontScaleSlider.Value;
                _accessibilityService.SpeechRate = (int)SpeechRateSlider.Value;
                _accessibilityService.HighContrastMode = HighContrastCheck.IsChecked ?? false;
                _accessibilityService.SaveSettings();
            }

            // アプリ設定を保存
            var selectedTheme = (ThemeComboBox.SelectedItem as ComboBoxItem)?.Tag?.ToString() ?? "ModernMinimal";

            var settings = new AppSettings
            {
                ClaudeApiKey = ApiKeyInput.Password,
                EmailAddress = EmailInput.Text,
                EmailPassword = EmailPasswordInput.Password,
                DisplayName = DisplayNameInput.Text,
                SmtpServer = SmtpServerInput.Text,
                SmtpPort = int.TryParse(SmtpPortInput.Text, out var smtpPort) ? smtpPort : 587,
                ImapServer = ImapServerInput.Text,
                ImapPort = int.TryParse(ImapPortInput.Text, out var imapPort) ? imapPort : 993,
                SpeechRate = (int)SpeechRateSlider.Value,
                ColorTheme = selectedTheme
            };

            var directory = Path.GetDirectoryName(SettingsFilePath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            var json = JsonSerializer.Serialize(settings, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(SettingsFilePath, json);
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"設定の保存に失敗しました。\n{ex.Message}",
                "エラー",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }

    private void OnApiKeyChanged(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(ApiKeyInput.Password))
        {
            ApiKeyStatus.Text = "";
        }
        else
        {
            ApiKeyStatus.Text = "APIキーが入力されています（保存後に有効になります）";
            ApiKeyStatus.Foreground = System.Windows.Media.Brushes.Orange;
        }
    }

    private void OnFontScaleChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
    {
        UpdateFontPreview();
    }

    private void UpdateFontPreview()
    {
        if (FontScalePercent == null || FontPreviewText == null || FontScaleSlider == null)
            return;

        var scale = FontScaleSlider.Value;
        var percent = (int)(scale * 100);
        FontScalePercent.Text = $"{percent}%";

        // プレビューテキストのサイズを更新
        FontPreviewText.FontSize = 22 * scale;
    }

    private void OnShowTutorialClick(object sender, RoutedEventArgs e)
    {
        var tutorialService = App.Services?.GetService<ITutorialService>();
        if (tutorialService != null)
        {
            tutorialService.Reset();
            var tutorialWindow = new TutorialWindow(tutorialService);
            tutorialWindow.Owner = this;
            tutorialWindow.ShowDialog();
        }
    }

    private void OnSaveClick(object sender, RoutedEventArgs e)
    {
        SaveSettings();

        MessageBox.Show(
            "設定を保存しました。",
            "保存完了",
            MessageBoxButton.OK,
            MessageBoxImage.Information);

        DialogResult = true;
        Close();
    }

    private void OnCancelClick(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }

    private void OnThemeChanged(object sender, SelectionChangedEventArgs e)
    {
        if (ThemeComboBox.SelectedItem is ComboBoxItem item && ThemePreview != null)
        {
            // プレビューを更新
            ThemePreview.Children.Clear();
            var theme = item.Tag?.ToString() ?? "ModernMinimal";

            var colors = GetThemeColors(theme);

            foreach (var (label, color) in colors)
            {
                var button = new Button
                {
                    Content = label,
                    Background = new System.Windows.Media.SolidColorBrush(
                        (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(color)),
                    Foreground = System.Windows.Media.Brushes.White,
                    FontSize = 12,
                    FontWeight = FontWeights.Bold,
                    Padding = new Thickness(10, 6, 10, 6),
                    Margin = new Thickness(4),
                    BorderThickness = new Thickness(0)
                };
                ThemePreview.Children.Add(button);
            }
        }
    }

    private List<(string Label, string Color)> GetThemeColors(string theme)
    {
        return theme switch
        {
            "GoogleSheets" => new List<(string, string)>
            {
                ("ファイル", "#F3F4F6"),
                ("合計", "#F3F4F6"),
                ("平均", "#F3F4F6"),
                ("消費税", "#F3F4F6"),
            },
            "SeniorFriendly" => new List<(string, string)>
            {
                ("ファイル", "#FFFFFF"),
                ("合計", "#FFFFFF"),
                ("平均", "#FFFFFF"),
                ("保存", "#3B82F6"),
            },
            "ModernMinimal" or _ => new List<(string, string)>
            {
                ("ファイル", "#F3F4F6"),
                ("合計", "#EFF6FF"),
                ("平均", "#EFF6FF"),
                ("保存", "#3B82F6"),
            }
        };
    }
}

/// <summary>
/// アプリケーション設定
/// </summary>
public class AppSettings
{
    public string? ClaudeApiKey { get; set; }
    public string? EmailAddress { get; set; }
    public string? EmailPassword { get; set; }
    public string? DisplayName { get; set; }
    public string? SmtpServer { get; set; }
    public int SmtpPort { get; set; } = 587;
    public string? ImapServer { get; set; }
    public int ImapPort { get; set; } = 993;
    public int SpeechRate { get; set; } = 0;
    public double FontScale { get; set; } = 1.0;
    public string ColorTheme { get; set; } = "ModernMinimal";
}
