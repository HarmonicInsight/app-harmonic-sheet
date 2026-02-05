using CommunityToolkit.Mvvm.ComponentModel;

namespace HarmonicSheet.Models;

/// <summary>
/// メールメッセージを表します
/// </summary>
public partial class MailMessage : ObservableObject
{
    [ObservableProperty]
    private string _id = string.Empty;

    [ObservableProperty]
    private string _from = string.Empty;

    [ObservableProperty]
    private string _to = string.Empty;

    [ObservableProperty]
    private string _subject = string.Empty;

    [ObservableProperty]
    private string _body = string.Empty;

    [ObservableProperty]
    private DateTime _date;

    [ObservableProperty]
    private bool _isRead;

    [ObservableProperty]
    private bool _hasAttachment;
}

/// <summary>
/// メール送信の設定
/// </summary>
public class MailSettings
{
    public string SmtpServer { get; set; } = string.Empty;
    public int SmtpPort { get; set; } = 587;
    public string ImapServer { get; set; } = string.Empty;
    public int ImapPort { get; set; } = 993;
    public string EmailAddress { get; set; } = string.Empty;
    public string Password { get; set; } = string.Empty;
    public string DisplayName { get; set; } = string.Empty;
}

/// <summary>
/// Claude APIからのメール操作結果
/// </summary>
public class MailCommandResult
{
    public bool Success { get; set; }
    public string Message { get; set; } = string.Empty;
    public string? Action { get; set; }  // "compose", "reply", "forward", etc.
    public string? To { get; set; }
    public string? Subject { get; set; }
    public string? Body { get; set; }
}
