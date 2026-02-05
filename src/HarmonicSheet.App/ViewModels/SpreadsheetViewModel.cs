using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using HarmonicSheet.Models;
using HarmonicSheet.Services;

namespace HarmonicSheet.ViewModels;

public partial class SpreadsheetViewModel : ObservableObject
{
    private readonly ISpreadsheetService _spreadsheetService;
    private readonly IClaudeService _claudeService;
    private readonly IPrintService _printService;

    [ObservableProperty]
    private ObservableCollection<SpreadsheetRow> _rows = new();

    [ObservableProperty]
    private string _commandInput = string.Empty;

    [ObservableProperty]
    private string _statusMessage = string.Empty;

    [ObservableProperty]
    private bool _isProcessing;

    public SpreadsheetViewModel(
        ISpreadsheetService spreadsheetService,
        IClaudeService claudeService,
        IPrintService printService)
    {
        _spreadsheetService = spreadsheetService;
        _claudeService = claudeService;
        _printService = printService;

        // 初期化: 10行5列の空のグリッドを作成
        InitializeGrid(10, 5);
    }

    private void InitializeGrid(int rowCount, int colCount)
    {
        Rows.Clear();
        for (int i = 0; i < rowCount; i++)
        {
            var row = new SpreadsheetRow { RowNumber = i + 1 };
            for (int j = 0; j < colCount; j++)
            {
                row.Cells.Add(new SpreadsheetCell
                {
                    Column = (char)('A' + j),
                    Row = i + 1,
                    Value = string.Empty
                });
            }
            Rows.Add(row);
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
            var result = await _claudeService.ProcessSpreadsheetCommand(CommandInput, GetCurrentGridState());

            if (result.Success)
            {
                // 結果を適用
                ApplyChanges(result.Changes);
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

    private string GetCurrentGridState()
    {
        var state = new System.Text.StringBuilder();
        foreach (var row in Rows)
        {
            foreach (var cell in row.Cells)
            {
                if (!string.IsNullOrEmpty(cell.Value))
                {
                    state.AppendLine($"{cell.Column}{cell.Row}: {cell.Value}");
                }
            }
        }
        return state.ToString();
    }

    private void ApplyChanges(List<CellChange> changes)
    {
        foreach (var change in changes)
        {
            var row = Rows.FirstOrDefault(r => r.RowNumber == change.Row);
            if (row != null)
            {
                var cell = row.Cells.FirstOrDefault(c => c.Column == change.Column);
                if (cell != null)
                {
                    cell.Value = change.Value;
                    cell.Formula = change.Formula;
                }
            }
        }
    }

    [RelayCommand]
    private void Print()
    {
        _printService.PrintSpreadsheet(Rows.ToList());
        StatusMessage = "印刷しました";
    }

    [RelayCommand]
    private async Task Save()
    {
        var dialog = new Microsoft.Win32.SaveFileDialog
        {
            Filter = "Excel ファイル (*.xlsx)|*.xlsx",
            DefaultExt = ".xlsx",
            FileName = "表データ"
        };

        if (dialog.ShowDialog() == true)
        {
            await _spreadsheetService.SaveToExcel(Rows.ToList(), dialog.FileName);
            StatusMessage = "保存しました";
        }
    }

    [RelayCommand]
    private async Task Open()
    {
        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Filter = "Excel ファイル (*.xlsx)|*.xlsx|すべてのファイル (*.*)|*.*",
            DefaultExt = ".xlsx"
        };

        if (dialog.ShowDialog() == true)
        {
            var loadedRows = await _spreadsheetService.LoadFromExcel(dialog.FileName);
            Rows.Clear();
            foreach (var row in loadedRows)
            {
                Rows.Add(row);
            }
            StatusMessage = "ファイルを開きました";
        }
    }
}
