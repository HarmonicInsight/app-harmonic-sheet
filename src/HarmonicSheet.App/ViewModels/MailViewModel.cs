using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using HarmonicSheet.Models;
using HarmonicSheet.Services;

namespace HarmonicSheet.ViewModels;

public partial class MailViewModel : ObservableObject
{
    private readonly IMailService _mailService;
    private readonly IClaudeService _claudeService;
    private readonly ISpeechService _speechService;
    private readonly IPrintService _printService;

    [ObservableProperty]
    private ObservableCollection<MailMessage> _messages = new();

    [ObservableProperty]
    private MailMessage? _selectedMessage;

    [ObservableProperty]
    private string _statusMessage = string.Empty;

    [ObservableProperty]
    private bool _isProcessing;

    // 新規メール作成用
    [ObservableProperty]
    private string _newMailTo = string.Empty;

    [ObservableProperty]
    private string _newMailSubject = string.Empty;

    [ObservableProperty]
    private string _newMailBody = string.Empty;

    [ObservableProperty]
    private string _commandInput = string.Empty;

    [ObservableProperty]
    private bool _isComposing;

    public MailViewModel(
        IMailService mailService,
        IClaudeService claudeService,
        ISpeechService speechService,
        IPrintService printService)
    {
        _mailService = mailService;
        _claudeService = claudeService;
        _speechService = speechService;
        _printService = printService;
    }

    [RelayCommand]
    private async Task RefreshMail()
    {
        IsProcessing = true;
        StatusMessage = "メールを取得中...";

        try
        {
            var mails = await _mailService.GetInboxMessages();
            Messages.Clear();
            foreach (var mail in mails)
            {
                Messages.Add(mail);
            }
            StatusMessage = $"{mails.Count}件のメールを取得しました";
        }
        catch (Exception ex)
        {
            StatusMessage = $"エラー: {ex.Message}";
        }
        finally
        {
            IsProcessing = false;
        }
    }

    [RelayCommand]
    private void StartCompose()
    {
        IsComposing = true;
        NewMailTo = string.Empty;
        NewMailSubject = string.Empty;
        NewMailBody = string.Empty;
        StatusMessage = "新しいメールを作成します";
    }

    [RelayCommand]
    private void CancelCompose()
    {
        IsComposing = false;
        StatusMessage = "メール作成をキャンセルしました";
    }

    [RelayCommand]
    private async Task SendMail()
    {
        if (string.IsNullOrWhiteSpace(NewMailTo))
        {
            StatusMessage = "宛先を入力してください";
            return;
        }

        IsProcessing = true;
        StatusMessage = "メールを送信中...";

        try
        {
            await _mailService.SendMessage(NewMailTo, NewMailSubject, NewMailBody);
            StatusMessage = "メールを送信しました";
            IsComposing = false;
        }
        catch (Exception ex)
        {
            StatusMessage = $"送信エラー: {ex.Message}";
        }
        finally
        {
            IsProcessing = false;
        }
    }

    [RelayCommand]
    private async Task ExecuteCommand()
    {
        if (string.IsNullOrWhiteSpace(CommandInput))
            return;

        IsProcessing = true;
        StatusMessage = "処理中...";

        try
        {
            // Claude Haikuに自然言語コマンドを送信
            var result = await _claudeService.ProcessMailCommand(CommandInput, SelectedMessage);

            if (result.Success)
            {
                // コマンド結果に応じた処理
                if (result.Action == "compose")
                {
                    IsComposing = true;
                    NewMailTo = result.To ?? string.Empty;
                    NewMailSubject = result.Subject ?? string.Empty;
                    NewMailBody = result.Body ?? string.Empty;
                }
                StatusMessage = result.Message;
            }
            else
            {
                StatusMessage = $"エラー: {result.Message}";
            }
        }
        catch (Exception ex)
        {
            StatusMessage = $"エラーが発生しました: {ex.Message}";
        }
        finally
        {
            IsProcessing = false;
            CommandInput = string.Empty;
        }
    }

    [RelayCommand]
    private void ReadSelectedMail()
    {
        if (SelectedMessage != null)
        {
            var textToRead = $"送信者: {SelectedMessage.From}。件名: {SelectedMessage.Subject}。本文: {SelectedMessage.Body}";
            _speechService.Speak(textToRead);
            StatusMessage = "メールを読み上げ中...";
        }
    }

    [RelayCommand]
    private void StopReading()
    {
        _speechService.StopSpeaking();
        StatusMessage = "読み上げを停止しました";
    }

    [RelayCommand]
    private void PrintSelectedMail()
    {
        if (SelectedMessage != null)
        {
            _printService.PrintMail(SelectedMessage);
            StatusMessage = "印刷しました";
        }
    }
}
