using MailKit.Net.Smtp;
using MailKit.Net.Imap;
using MailKit.Security;
using MimeKit;
using HarmonicSheet.Models;

namespace HarmonicSheet.Services;

/// <summary>
/// メールサービスの実装（MailKit使用）
/// </summary>
public class MailService : IMailService
{
    private MailSettings? _settings;

    public bool IsConfigured => _settings != null &&
        !string.IsNullOrEmpty(_settings.SmtpServer) &&
        !string.IsNullOrEmpty(_settings.EmailAddress);

    public void Configure(MailSettings settings)
    {
        _settings = settings;
    }

    public async Task<List<MailMessage>> GetInboxMessages(int maxCount = 20)
    {
        if (!IsConfigured || _settings == null)
        {
            throw new InvalidOperationException("メール設定がされていません。設定画面でメールアカウントを設定してください。");
        }

        var messages = new List<MailMessage>();

        using var client = new ImapClient();

        await client.ConnectAsync(_settings.ImapServer, _settings.ImapPort, SecureSocketOptions.SslOnConnect);
        await client.AuthenticateAsync(_settings.EmailAddress, _settings.Password);

        var inbox = client.Inbox;
        await inbox.OpenAsync(MailKit.FolderAccess.ReadOnly);

        // 最新のメールから取得
        var startIndex = Math.Max(0, inbox.Count - maxCount);
        for (int i = inbox.Count - 1; i >= startIndex; i--)
        {
            var message = await inbox.GetMessageAsync(i);

            messages.Add(new MailMessage
            {
                Id = message.MessageId ?? Guid.NewGuid().ToString(),
                From = message.From.ToString(),
                To = message.To.ToString(),
                Subject = message.Subject ?? "(件名なし)",
                Body = message.TextBody ?? message.HtmlBody ?? "",
                Date = message.Date.DateTime,
                IsRead = true, // 簡略化のため
                HasAttachment = message.Attachments.Any()
            });
        }

        await client.DisconnectAsync(true);

        return messages;
    }

    public async Task SendMessage(string to, string subject, string body)
    {
        if (!IsConfigured || _settings == null)
        {
            throw new InvalidOperationException("メール設定がされていません。設定画面でメールアカウントを設定してください。");
        }

        var message = new MimeMessage();
        message.From.Add(new MailboxAddress(_settings.DisplayName, _settings.EmailAddress));
        message.To.Add(MailboxAddress.Parse(to));
        message.Subject = subject;

        message.Body = new TextPart("plain")
        {
            Text = body
        };

        using var client = new SmtpClient();

        await client.ConnectAsync(_settings.SmtpServer, _settings.SmtpPort, SecureSocketOptions.StartTls);
        await client.AuthenticateAsync(_settings.EmailAddress, _settings.Password);
        await client.SendAsync(message);
        await client.DisconnectAsync(true);
    }
}
