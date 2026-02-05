using HarmonicSheet.Models;

namespace HarmonicSheet.Services;

/// <summary>
/// メールサービスのインターフェース
/// </summary>
public interface IMailService
{
    /// <summary>
    /// メール設定を構成します
    /// </summary>
    void Configure(MailSettings settings);

    /// <summary>
    /// 設定済みかどうか
    /// </summary>
    bool IsConfigured { get; }

    /// <summary>
    /// 受信トレイのメールを取得します
    /// </summary>
    Task<List<MailMessage>> GetInboxMessages(int maxCount = 20);

    /// <summary>
    /// メールを送信します
    /// </summary>
    Task SendMessage(string to, string subject, string body);
}
