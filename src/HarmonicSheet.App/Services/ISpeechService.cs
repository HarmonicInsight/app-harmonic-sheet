namespace HarmonicSheet.Services;

/// <summary>
/// 音声入力・読み上げサービスのインターフェース
/// </summary>
public interface ISpeechService
{
    /// <summary>
    /// テキストを読み上げます
    /// </summary>
    void Speak(string text);

    /// <summary>
    /// 読み上げを停止します
    /// </summary>
    void StopSpeaking();

    /// <summary>
    /// Windows音声入力（Win+H）を起動します
    /// </summary>
    void ActivateWindowsVoiceTyping();

    /// <summary>
    /// 読み上げ中かどうか
    /// </summary>
    bool IsSpeaking { get; }

    /// <summary>
    /// 読み上げ速度（0-10、デフォルト5）
    /// </summary>
    int SpeechRate { get; set; }
}
