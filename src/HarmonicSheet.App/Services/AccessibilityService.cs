using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text.Json;
using System.Windows;

namespace HarmonicSheet.Services;

/// <summary>
/// アクセシビリティ設定サービス
/// フォントサイズ、音声速度などを管理
/// </summary>
public interface IAccessibilityService : INotifyPropertyChanged
{
    /// <summary>フォントサイズスケール（1.0 = 100%、1.5 = 150%）</summary>
    double FontScale { get; set; }

    /// <summary>ベースフォントサイズ（スケール適用前）</summary>
    double BaseFontSize { get; }

    /// <summary>実際のフォントサイズ（スケール適用後）</summary>
    double ActualFontSize { get; }

    /// <summary>大きなフォントサイズ</summary>
    double LargeFontSize { get; }

    /// <summary>特大フォントサイズ</summary>
    double HugeFontSize { get; }

    /// <summary>音声読み上げ速度（-10 ～ 10）</summary>
    int SpeechRate { get; set; }

    /// <summary>高コントラストモード</summary>
    bool HighContrastMode { get; set; }

    /// <summary>設定を保存</summary>
    void SaveSettings();

    /// <summary>設定をリセット</summary>
    void ResetToDefault();
}

public class AccessibilityService : IAccessibilityService
{
    private static readonly string SettingsFilePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "HarmonicSheet",
        "accessibility.json");

    private double _fontScale = 1.0;
    private int _speechRate = 0;
    private bool _highContrastMode = false;

    private const double BaseFont = 22.0;  // シニア向けのベースサイズ

    public event PropertyChangedEventHandler? PropertyChanged;

    public AccessibilityService()
    {
        LoadSettings();
    }

    public double FontScale
    {
        get => _fontScale;
        set
        {
            if (Math.Abs(_fontScale - value) > 0.01)
            {
                _fontScale = Math.Clamp(value, 0.5, 2.0);
                OnPropertyChanged();
                OnPropertyChanged(nameof(ActualFontSize));
                OnPropertyChanged(nameof(LargeFontSize));
                OnPropertyChanged(nameof(HugeFontSize));
                UpdateApplicationResources();
            }
        }
    }

    public double BaseFontSize => BaseFont;
    public double ActualFontSize => BaseFont * _fontScale;
    public double LargeFontSize => 28 * _fontScale;
    public double HugeFontSize => 48 * _fontScale;

    public int SpeechRate
    {
        get => _speechRate;
        set
        {
            if (_speechRate != value)
            {
                _speechRate = Math.Clamp(value, -10, 10);
                OnPropertyChanged();
            }
        }
    }

    public bool HighContrastMode
    {
        get => _highContrastMode;
        set
        {
            if (_highContrastMode != value)
            {
                _highContrastMode = value;
                OnPropertyChanged();
                UpdateApplicationResources();
            }
        }
    }

    private void LoadSettings()
    {
        try
        {
            if (File.Exists(SettingsFilePath))
            {
                var json = File.ReadAllText(SettingsFilePath);
                var settings = JsonSerializer.Deserialize<AccessibilitySettings>(json);
                if (settings != null)
                {
                    _fontScale = settings.FontScale;
                    _speechRate = settings.SpeechRate;
                    _highContrastMode = settings.HighContrastMode;
                }
            }
        }
        catch
        {
            // 読み込みエラーは無視
        }
    }

    public void SaveSettings()
    {
        try
        {
            var directory = Path.GetDirectoryName(SettingsFilePath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            var settings = new AccessibilitySettings
            {
                FontScale = _fontScale,
                SpeechRate = _speechRate,
                HighContrastMode = _highContrastMode
            };

            var json = JsonSerializer.Serialize(settings, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(SettingsFilePath, json);
        }
        catch
        {
            // 保存エラーは無視
        }
    }

    public void ResetToDefault()
    {
        FontScale = 1.0;
        SpeechRate = 0;
        HighContrastMode = false;
        SaveSettings();
    }

    private void UpdateApplicationResources()
    {
        // アプリケーション全体のリソースを更新
        if (Application.Current?.Resources != null)
        {
            Application.Current.Resources["FontSizeSmall"] = 18.0 * _fontScale;
            Application.Current.Resources["FontSizeNormal"] = ActualFontSize;
            Application.Current.Resources["FontSizeLarge"] = LargeFontSize;
            Application.Current.Resources["FontSizeXLarge"] = 36.0 * _fontScale;
            Application.Current.Resources["FontSizeHuge"] = HugeFontSize;
        }
    }

    protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    private class AccessibilitySettings
    {
        public double FontScale { get; set; } = 1.0;
        public int SpeechRate { get; set; } = 0;
        public bool HighContrastMode { get; set; } = false;
    }
}
