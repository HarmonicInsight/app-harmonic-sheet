using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using HarmonicSheet.Services;
using HarmonicSheet.ViewModels;
using Microsoft.Win32;

namespace HarmonicSheet.Views;

public partial class SpreadsheetView : UserControl
{
    private readonly IClaudeService? _claudeService;
    private readonly IPrintService? _printService;

    public SpreadsheetView()
    {
        InitializeComponent();

        // DIが使える場合はサービスを取得
        if (Application.Current is App app)
        {
            // サービスは後で設定される
        }

        // 新規ワークブックを作成
        Spreadsheet.Create(10, 5);

        // シニア向けの大きなフォント設定
        ConfigureForSeniors();
    }

    private void ConfigureForSeniors()
    {
        // 行高さと列幅を大きく設定
        try
        {
            var worksheet = Spreadsheet.ActiveSheet;
            if (worksheet != null)
            {
                // 全ての行の高さを大きく
                for (int i = 1; i <= 100; i++)
                {
                    worksheet.SetRowHeight(i, 30);
                }
                // 全ての列の幅を大きく
                for (int i = 1; i <= 26; i++)
                {
                    worksheet.SetColumnWidth(i, 120);
                }
            }
        }
        catch
        {
            // 設定に失敗しても続行
        }
    }

    private void OnNewClick(object sender, RoutedEventArgs e)
    {
        var result = MessageBox.Show(
            "新しい表を作りますか？\n今の内容は消えます。",
            "確認",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question);

        if (result == MessageBoxResult.Yes)
        {
            Spreadsheet.Create(10, 5);
            ConfigureForSeniors();
            StatusText.Text = "新しい表を作りました";
        }
    }

    private void OnOpenClick(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Excel ファイル (*.xlsx)|*.xlsx|すべてのファイル (*.*)|*.*",
            Title = "開くファイルを選んでください"
        };

        if (dialog.ShowDialog() == true)
        {
            try
            {
                Spreadsheet.Open(dialog.FileName);
                StatusText.Text = $"ファイルを開きました: {System.IO.Path.GetFileName(dialog.FileName)}";
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"ファイルを開けませんでした。\n{ex.Message}",
                    "エラー",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }
    }

    private void OnSaveClick(object sender, RoutedEventArgs e)
    {
        var dialog = new SaveFileDialog
        {
            Filter = "Excel ファイル (*.xlsx)|*.xlsx",
            Title = "保存するファイル名を入力してください",
            FileName = "表データ"
        };

        if (dialog.ShowDialog() == true)
        {
            try
            {
                Spreadsheet.Save(dialog.FileName);
                StatusText.Text = $"保存しました: {System.IO.Path.GetFileName(dialog.FileName)}";
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"保存できませんでした。\n{ex.Message}",
                    "エラー",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }
    }

    private void OnPrintClick(object sender, RoutedEventArgs e)
    {
        try
        {
            Spreadsheet.Print();
            StatusText.Text = "印刷しました";
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"印刷できませんでした。\n{ex.Message}",
                "エラー",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }

    private void OnCommandKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter)
        {
            OnExecuteCommand(sender, e);
        }
    }

    private async void OnExecuteCommand(object sender, RoutedEventArgs e)
    {
        var command = CommandInput.Text?.Trim();
        if (string.IsNullOrEmpty(command))
        {
            StatusText.Text = "やりたいことを入力してください";
            return;
        }

        StatusText.Text = "処理中...";

        try
        {
            // 簡易的な自然言語パース（Claude APIが使えない場合のフォールバック）
            var result = ParseAndExecuteCommand(command);
            StatusText.Text = result;
            CommandInput.Text = string.Empty;
        }
        catch (Exception ex)
        {
            StatusText.Text = $"エラー: {ex.Message}";
        }
    }

    private string ParseAndExecuteCommand(string command)
    {
        // 簡易的なコマンドパーサー
        // 「A2に1万円入れて」「A1+A2+A3をA4に」などを解析

        try
        {
            // パターン1: 「〇〇に△△を入れて」
            if (command.Contains("入れ") || command.Contains("いれ"))
            {
                return ExecuteSetValueCommand(command);
            }

            // パターン2: 「〇〇を足して△△に」
            if (command.Contains("足") || command.Contains("合計") || command.Contains("プラス"))
            {
                return ExecuteSumCommand(command);
            }

            return "すみません、よくわかりませんでした。もう一度教えてください。";
        }
        catch
        {
            return "すみません、うまくいきませんでした。";
        }
    }

    private string ExecuteSetValueCommand(string command)
    {
        // 「A2に1万円入れて」を解析
        // セルアドレスを探す
        var cellMatch = System.Text.RegularExpressions.Regex.Match(
            command,
            @"([A-Za-z])[\s]*[のノ]?[\s]*(\d+)|([A-Za-z])(\d+)");

        if (!cellMatch.Success)
        {
            return "セルの場所がわかりませんでした。（例：A2、Aの2）";
        }

        string col, row;
        if (!string.IsNullOrEmpty(cellMatch.Groups[1].Value))
        {
            col = cellMatch.Groups[1].Value.ToUpper();
            row = cellMatch.Groups[2].Value;
        }
        else
        {
            col = cellMatch.Groups[3].Value.ToUpper();
            row = cellMatch.Groups[4].Value;
        }

        // 数値を探す
        var valueMatch = System.Text.RegularExpressions.Regex.Match(
            command,
            @"(\d+)[万]?[円]?|(\d+,?\d*)");

        if (!valueMatch.Success)
        {
            return "入れる値がわかりませんでした。";
        }

        var valueStr = valueMatch.Groups[1].Success ? valueMatch.Groups[1].Value : valueMatch.Groups[2].Value;
        valueStr = valueStr.Replace(",", "");

        // 「万」があれば10000倍
        if (command.Contains("万"))
        {
            if (double.TryParse(valueStr, out var num))
            {
                valueStr = (num * 10000).ToString();
            }
        }

        // セルに値を設定
        var worksheet = Spreadsheet.ActiveSheet;
        var colIndex = col[0] - 'A' + 1;
        var rowIndex = int.Parse(row);

        worksheet[$"{col}{row}"].Value = valueStr;
        Spreadsheet.ActiveGrid.InvalidateCell(rowIndex, colIndex);

        return $"{col}{row} に {valueStr} を入れました";
    }

    private string ExecuteSumCommand(string command)
    {
        // 「A1とA2とA3を足してA4に」を解析
        var cellMatches = System.Text.RegularExpressions.Regex.Matches(
            command,
            @"([A-Za-z])[\s]*[のノ]?[\s]*(\d+)|([A-Za-z])(\d+)");

        if (cellMatches.Count < 2)
        {
            return "計算するセルと結果を入れるセルを教えてください。";
        }

        var cells = new List<string>();
        foreach (System.Text.RegularExpressions.Match match in cellMatches)
        {
            string col, row;
            if (!string.IsNullOrEmpty(match.Groups[1].Value))
            {
                col = match.Groups[1].Value.ToUpper();
                row = match.Groups[2].Value;
            }
            else
            {
                col = match.Groups[3].Value.ToUpper();
                row = match.Groups[4].Value;
            }
            cells.Add($"{col}{row}");
        }

        // 最後のセルが結果を入れる場所
        var targetCell = cells[cells.Count - 1];
        var sourceCells = cells.Take(cells.Count - 1).ToList();

        // 数式を作成
        var formula = "=" + string.Join("+", sourceCells);

        // セルに数式を設定
        var worksheet = Spreadsheet.ActiveSheet;
        worksheet[targetCell].Formula = formula;

        var colIndex = targetCell[0] - 'A' + 1;
        var rowIndex = int.Parse(targetCell.Substring(1));
        Spreadsheet.ActiveGrid.InvalidateCell(rowIndex, colIndex);

        return $"{string.Join("と", sourceCells)} を足して {targetCell} に入れました";
    }
}
