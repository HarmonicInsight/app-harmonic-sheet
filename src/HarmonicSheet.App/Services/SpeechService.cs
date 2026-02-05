using System.Speech.Synthesis;
using System.Runtime.InteropServices;

namespace HarmonicSheet.Services;

/// <summary>
/// 音声入力・読み上げサービスの実装
/// </summary>
public class SpeechService : ISpeechService, IDisposable
{
    private readonly SpeechSynthesizer _synthesizer;
    private bool _disposed;

    public SpeechService()
    {
        _synthesizer = new SpeechSynthesizer();

        // 日本語音声を設定（利用可能な場合）
        try
        {
            var japaneseVoice = _synthesizer.GetInstalledVoices()
                .FirstOrDefault(v => v.VoiceInfo.Culture.Name.StartsWith("ja"));
            if (japaneseVoice != null)
            {
                _synthesizer.SelectVoice(japaneseVoice.VoiceInfo.Name);
            }
        }
        catch
        {
            // 日本語音声が見つからない場合はデフォルトを使用
        }

        // シニア向けにゆっくり読み上げ
        _synthesizer.Rate = 0; // -10 (遅い) から 10 (速い) の範囲、0が中間
    }

    public bool IsSpeaking => _synthesizer.State == SynthesizerState.Speaking;

    public int SpeechRate
    {
        get => _synthesizer.Rate + 5; // 0-10の範囲に変換
        set => _synthesizer.Rate = Math.Clamp(value - 5, -10, 10);
    }

    public void Speak(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return;

        // 既に読み上げ中なら停止
        if (IsSpeaking)
        {
            StopSpeaking();
        }

        // 非同期で読み上げ
        _synthesizer.SpeakAsync(text);
    }

    public void StopSpeaking()
    {
        if (IsSpeaking)
        {
            _synthesizer.SpeakAsyncCancelAll();
        }
    }

    public void ActivateWindowsVoiceTyping()
    {
        // Win+H キーストロークをシミュレート
        // Windows 10/11の音声入力を起動
        try
        {
            // Win キーを押す
            keybd_event(VK_LWIN, 0, 0, IntPtr.Zero);
            // H キーを押す
            keybd_event(VK_H, 0, 0, IntPtr.Zero);
            // H キーを離す
            keybd_event(VK_H, 0, KEYEVENTF_KEYUP, IntPtr.Zero);
            // Win キーを離す
            keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, IntPtr.Zero);
        }
        catch
        {
            // キー入力のシミュレートに失敗した場合は無視
        }
    }

    // Win32 API for keyboard simulation
    [DllImport("user32.dll", SetLastError = true)]
    private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, IntPtr dwExtraInfo);

    private const byte VK_LWIN = 0x5B;  // Left Windows key
    private const byte VK_H = 0x48;      // H key
    private const uint KEYEVENTF_KEYUP = 0x0002;

    public void Dispose()
    {
        if (!_disposed)
        {
            StopSpeaking();
            _synthesizer.Dispose();
            _disposed = true;
        }
    }
}
