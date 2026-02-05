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
                SpeechRate = (int)SpeechRateSlider.Value
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
}
