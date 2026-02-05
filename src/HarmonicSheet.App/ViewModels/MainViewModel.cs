using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using HarmonicSheet.Services;

namespace HarmonicSheet.ViewModels;

public partial class MainViewModel : ObservableObject
{
    private readonly ISpeechService _speechService;
    private readonly IClaudeService _claudeService;

    [ObservableProperty]
    private string _statusMessage = "準備完了";

    [ObservableProperty]
    private bool _isProcessing;

    public MainViewModel(ISpeechService speechService, IClaudeService claudeService)
    {
        _speechService = speechService;
        _claudeService = claudeService;
    }

    [RelayCommand]
    private void StartVoiceInput()
    {
        _speechService.ActivateWindowsVoiceTyping();
        StatusMessage = "音声入力中...";
    }
}
