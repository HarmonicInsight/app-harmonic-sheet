using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using HarmonicSheet.Models;
using HarmonicSheet.Services;

namespace HarmonicSheet.Views;

public partial class MailView : UserControl
{
    private readonly ObservableCollection<MailMessage> _messages = new();
    private MailMessage? _selectedMail;

    public MailView()
    {
        InitializeComponent();
        MailList.ItemsSource = _messages;

        // デモ用のサンプルメール
        LoadSampleMails();
    }

    private void LoadSampleMails()
    {
        _messages.Add(new MailMessage
        {
            Id = "1",
            From = "田中太郎 <tanaka@example.com>",
            Subject = "お元気ですか？",
            Body = "こんにちは。\n\nお元気ですか？\n先日はありがとうございました。\n\nまた遊びに来てくださいね。\n\n田中太郎",
            Date = DateTime.Now.AddDays(-1)
        });

        _messages.Add(new MailMessage
        {
            Id = "2",
            From = "山田花子 <yamada@example.com>",
            Subject = "来週の予定について",
            Body = "来週の土曜日、一緒にお食事いかがですか？\n\n近くのレストランを予約しておきます。\n\n山田花子",
            Date = DateTime.Now.AddDays(-2)
        });

        _messages.Add(new MailMessage
        {
            Id = "3",
            From = "病院 <hospital@example.com>",
            Subject = "次回の診察予約のお知らせ",
            Body = "次回の診察予約をお知らせします。\n\n日時：来週水曜日 午後2時\n場所：内科\n\nお忘れなくお越しください。",
            Date = DateTime.Now.AddDays(-3)
        });
    }

    private void OnMailSelected(object sender, SelectionChangedEventArgs e)
    {
        if (MailList.SelectedItem is MailMessage mail)
        {
            _selectedMail = mail;
            MailFrom.Text = $"送信者: {mail.From}";
            MailDate.Text = $"日時: {mail.Date:yyyy年M月d日 H:mm}";
            MailSubject.Text = mail.Subject;
            MailBody.Text = mail.Body;
            StatusText.Text = "メールを表示しています";
        }
    }

    private void OnNewMailClick(object sender, RoutedEventArgs e)
    {
        ShowComposePanel();
        ToInput.Text = string.Empty;
        SubjectInput.Text = string.Empty;
        BodyInput.Text = string.Empty;
        StatusText.Text = "新しいメールを作成";
    }

    private void OnRefreshClick(object sender, RoutedEventArgs e)
    {
        StatusText.Text = "メールを受信中...";

        // 実際にはメールサーバーから取得
        // ここではデモなのでメッセージのみ
        MessageBox.Show(
            "メールを受信するには、設定画面で\nメールアカウントを設定してください。",
            "メール設定",
            MessageBoxButton.OK,
            MessageBoxImage.Information);

        StatusText.Text = "メール設定が必要です";
    }

    private void OnReplyClick(object sender, RoutedEventArgs e)
    {
        if (_selectedMail == null)
        {
            StatusText.Text = "返信するメールを選んでください";
            return;
        }

        ShowComposePanel();
        ToInput.Text = _selectedMail.From;
        SubjectInput.Text = $"Re: {_selectedMail.Subject}";
        BodyInput.Text = $"\n\n---元のメール---\n{_selectedMail.Body}";
        StatusText.Text = "返信を作成";
    }

    private void OnReadAloudClick(object sender, RoutedEventArgs e)
    {
        if (_selectedMail == null)
        {
            StatusText.Text = "読み上げるメールを選んでください";
            return;
        }

        try
        {
            var synthesizer = new System.Speech.Synthesis.SpeechSynthesizer();

            // 日本語音声を探す
            var japaneseVoice = synthesizer.GetInstalledVoices()
                .FirstOrDefault(v => v.VoiceInfo.Culture.Name.StartsWith("ja"));
            if (japaneseVoice != null)
            {
                synthesizer.SelectVoice(japaneseVoice.VoiceInfo.Name);
            }

            synthesizer.Rate = 0;

            var textToRead = $"送信者、{_selectedMail.From}。件名、{_selectedMail.Subject}。本文、{_selectedMail.Body}";
            synthesizer.SpeakAsync(textToRead);

            StatusText.Text = "読み上げ中...";
        }
        catch (Exception ex)
        {
            StatusText.Text = $"読み上げできませんでした: {ex.Message}";
        }
    }

    private void OnPrintMailClick(object sender, RoutedEventArgs e)
    {
        if (_selectedMail == null)
        {
            StatusText.Text = "印刷するメールを選んでください";
            return;
        }

        MessageBox.Show(
            "印刷機能は準備中です。",
            "印刷",
            MessageBoxButton.OK,
            MessageBoxImage.Information);
    }

    private void ShowComposePanel()
    {
        ReadPanel.Visibility = Visibility.Collapsed;
        ComposePanel.Visibility = Visibility.Visible;
    }

    private void ShowReadPanel()
    {
        ComposePanel.Visibility = Visibility.Collapsed;
        ReadPanel.Visibility = Visibility.Visible;
    }

    private void OnCancelComposeClick(object sender, RoutedEventArgs e)
    {
        ShowReadPanel();
        StatusText.Text = "メール作成をキャンセルしました";
    }

    private async void OnSendMailClick(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(ToInput.Text))
        {
            StatusText.Text = "宛先を入力してください";
            return;
        }

        StatusText.Text = "送信中...";

        await Task.Delay(1000); // デモ用の遅延

        MessageBox.Show(
            "メールを送信するには、設定画面で\nメールアカウントを設定してください。",
            "メール設定",
            MessageBoxButton.OK,
            MessageBoxImage.Information);

        StatusText.Text = "メール設定が必要です";
    }

    private async void OnAiAssistClick(object sender, RoutedEventArgs e)
    {
        var command = AiCommandInput.Text?.Trim();
        if (string.IsNullOrEmpty(command))
        {
            StatusText.Text = "AIにお願いしたいことを入力してください";
            return;
        }

        StatusText.Text = "AIに依頼中...";

        // デモ用の簡易応答
        await Task.Delay(500);

        if (command.Contains("お礼") || command.Contains("ありがとう"))
        {
            BodyInput.Text = "このたびは誠にありがとうございました。\n\nおかげさまで大変助かりました。\n今後ともよろしくお願いいたします。\n\n敬具";
            StatusText.Text = "お礼のメールを作成しました";
        }
        else if (command.Contains("挨拶") || command.Contains("お元気"))
        {
            BodyInput.Text = "お元気ですか。\n\nこちらは元気に過ごしております。\nまたお会いできることを楽しみにしています。\n\n敬具";
            StatusText.Text = "挨拶のメールを作成しました";
        }
        else
        {
            StatusText.Text = "AIアシスタントは設定画面でAPIキーを入力すると使えます";
        }

        AiCommandInput.Text = string.Empty;
    }
}
