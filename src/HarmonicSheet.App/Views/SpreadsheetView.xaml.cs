using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using HarmonicSheet.Services;
using HarmonicSheet.ViewModels;
using Microsoft.Win32;

namespace HarmonicSheet.Views;

public partial class SpreadsheetView : UserControl
{
    private static readonly string RecentFilesPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "HarmonicSheet",
        "recent_spreadsheets.txt");

    public SpreadsheetView()
    {
        InitializeComponent();

        // DIが使える場合はサービスを取得
        if (Application.Current is App app)
        {
            // サービスは後で設定される
        }

        // 新規ワークブックを作成（1シート）
        try
        {
            Spreadsheet.Create(1);

            // シニア向けの大きなフォント設定
            ConfigureForSeniors();
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Spreadsheet initialization error: {ex.Message}");
            MessageBox.Show($"表の初期化に失敗しました: {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
        }
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

    private void OnRecentClick(object sender, RoutedEventArgs e)
    {
        var recentFiles = LoadRecentFiles();
        if (recentFiles.Count == 0)
        {
            MessageBox.Show("最近使用したファイルはありません。", "前回使用", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        // 簡易的なファイル選択ダイアログ
        var message = "最近使用したファイル:\n\n";
        for (int i = 0; i < Math.Min(5, recentFiles.Count); i++)
        {
            message += $"{i + 1}. {Path.GetFileName(recentFiles[i])}\n";
        }
        message += "\n最新のファイルを開きますか？";

        var result = MessageBox.Show(message, "前回使用", MessageBoxButton.YesNo, MessageBoxImage.Question);
        if (result == MessageBoxResult.Yes && File.Exists(recentFiles[0]))
        {
            try
            {
                Spreadsheet.Open(recentFiles[0]);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ファイルを開けませんでした。\n{ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
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
            Spreadsheet.Create(1);
            ConfigureForSeniors();
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
                AddToRecentFiles(dialog.FileName);
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
                Spreadsheet.SaveAs(dialog.FileName);
                AddToRecentFiles(dialog.FileName);
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
            // SfSpreadsheet doesn't have a direct Print method
            // For now, show a message that printing requires saving to Excel first
            MessageBox.Show(
                "表の印刷をするには、一度Excelファイルとして保存してから、Excelで開いて印刷してください。",
                "印刷について",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
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
        if (e.Key == Key.Enter && !Keyboard.Modifiers.HasFlag(ModifierKeys.Shift))
        {
            OnExecuteCommand(sender, e);
            e.Handled = true;
        }
    }

    private void OnExecuteCommand(object sender, RoutedEventArgs e)
    {
        var command = CommandInput.Text?.Trim();
        if (string.IsNullOrEmpty(command))
        {
            return;
        }

        try
        {
            // 簡易的な自然言語パース（Claude APIが使えない場合のフォールバック）
            var result = ParseAndExecuteCommand(command);
            MessageBox.Show(result, "実行結果", MessageBoxButton.OK, MessageBoxImage.Information);
            CommandInput.Text = string.Empty;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"エラー: {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
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

    private List<string> LoadRecentFiles()
    {
        try
        {
            if (File.Exists(RecentFilesPath))
            {
                return File.ReadAllLines(RecentFilesPath)
                    .Where(f => File.Exists(f))
                    .ToList();
            }
        }
        catch { }
        return new List<string>();
    }

    private void AddToRecentFiles(string filePath)
    {
        try
        {
            var recentFiles = LoadRecentFiles();
            recentFiles.Remove(filePath); // 既存を削除
            recentFiles.Insert(0, filePath); // 先頭に追加

            // 最大10件まで保持
            var filesToSave = recentFiles.Take(10).ToList();

            var directory = Path.GetDirectoryName(RecentFilesPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            File.WriteAllLines(RecentFilesPath, filesToSave);
        }
        catch { }
    }
}
