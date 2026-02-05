using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using HarmonicSheet.Services;

namespace HarmonicSheet.ViewModels;

public partial class DocumentViewModel : ObservableObject
{
    private readonly IDocumentService _documentService;
    private readonly IClaudeService _claudeService;
    private readonly IPrintService _printService;
    private readonly ISpeechService _speechService;

    [ObservableProperty]
    private string _documentText = string.Empty;

    [ObservableProperty]
    private string _commandInput = string.Empty;

    [ObservableProperty]
    private string _statusMessage = string.Empty;

    [ObservableProperty]
    private bool _isProcessing;

    [ObservableProperty]
    private string _currentFileName = "新しい文書";

    public DocumentViewModel(
        IDocumentService documentService,
        IClaudeService claudeService,
        IPrintService printService,
        ISpeechService speechService)
    {
        _documentService = documentService;
        _claudeService = claudeService;
        _printService = printService;
        _speechService = speechService;
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
            var result = await _claudeService.ProcessDocumentCommand(CommandInput, DocumentText);

            if (result.Success)
            {
                if (!string.IsNullOrEmpty(result.NewText))
                {
                    DocumentText = result.NewText;
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
    private void ReadAloud()
    {
        if (!string.IsNullOrEmpty(DocumentText))
        {
            _speechService.Speak(DocumentText);
            StatusMessage = "読み上げ中...";
        }
    }

    [RelayCommand]
    private void StopReading()
    {
        _speechService.StopSpeaking();
        StatusMessage = "読み上げを停止しました";
    }

    [RelayCommand]
    private void Print()
    {
        _printService.PrintDocument(DocumentText, CurrentFileName);
        StatusMessage = "印刷しました";
    }

    [RelayCommand]
    private async Task Save()
    {
        var dialog = new Microsoft.Win32.SaveFileDialog
        {
            Filter = "テキストファイル (*.txt)|*.txt|すべてのファイル (*.*)|*.*",
            DefaultExt = ".txt",
            FileName = CurrentFileName
        };

        if (dialog.ShowDialog() == true)
        {
            await _documentService.SaveDocument(DocumentText, dialog.FileName);
            CurrentFileName = System.IO.Path.GetFileName(dialog.FileName);
            StatusMessage = "保存しました";
        }
    }

    [RelayCommand]
    private async Task Open()
    {
        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Filter = "テキストファイル (*.txt)|*.txt|すべてのファイル (*.*)|*.*",
            DefaultExt = ".txt"
        };

        if (dialog.ShowDialog() == true)
        {
            DocumentText = await _documentService.LoadDocument(dialog.FileName);
            CurrentFileName = System.IO.Path.GetFileName(dialog.FileName);
            StatusMessage = "ファイルを開きました";
        }
    }

    [RelayCommand]
    private void NewDocument()
    {
        DocumentText = string.Empty;
        CurrentFileName = "新しい文書";
        StatusMessage = "新しい文書を作成しました";
    }
}
