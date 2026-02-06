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
            var worksheet = Spreadsheet.ActiveSheet;
            if (worksheet == null)
            {
                MessageBox.Show("印刷するシートがありません。", "印刷", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // 印刷ダイアログを表示
            var printDialog = new System.Windows.Controls.PrintDialog();
            if (printDialog.ShowDialog() == true)
            {
                try
                {
                    // Workbookを取得してPrintSettingsを設定
                    var workbook = Spreadsheet.Workbook;
                    if (workbook != null && workbook.Worksheets.Count > 0)
                    {
                        // 全てのワークシートに印刷設定を適用
                        foreach (var ws in workbook.Worksheets)
                        {
                            ws.PageSetup.FitToPagesTall = 1;
                            ws.PageSetup.FitToPagesWide = 1;
                            ws.PageSetup.IsFitToPage = true;
                        }
                    }

                    // 一時ファイルに保存して印刷
                    var tempFile = Path.Combine(Path.GetTempPath(), $"print_{DateTime.Now:yyyyMMddHHmmss}.xlsx");
                    Spreadsheet.SaveAs(tempFile);

                    MessageBox.Show(
                        $"印刷プレビューを準備しました。\n\n保存先: {tempFile}\n\nこのファイルをExcelで開いて印刷してください。\n設定: 全体を1ページに収める設定済み",
                        "印刷準備完了",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);

                    // ファイルを開く
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = tempFile,
                        UseShellExecute = true
                    });
                }
                catch
                {
                    MessageBox.Show("印刷プレビューを表示できませんでした。", "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"印刷できませんでした。\nエラー: {ex.Message}\n\n別の方法として、「保存」ボタンでExcelファイルとして保存してから印刷することもできます。",
                "印刷エラー",
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

    // ========================================
    // 計算ボタンのハンドラー
    // ========================================

    private void OnSumClick(object sender, RoutedEventArgs e)
    {
        try
        {
            var selectedRanges = Spreadsheet.ActiveGrid.SelectedRanges;
            if (selectedRanges == null || selectedRanges.Count == 0)
            {
                MessageBox.Show("セルを選択してから合計ボタンを押してください。", "合計", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var range = selectedRanges[0];
            var worksheet = Spreadsheet.ActiveSheet;

            // 選択範囲の下または右に結果を出力
            var targetRow = range.Bottom + 1;
            var targetCol = range.Left;

            var rangeAddress = $"{GetColumnName(range.Left)}{range.Top}:{GetColumnName(range.Right)}{range.Bottom}";
            var formula = $"=SUM({rangeAddress})";

            worksheet[$"{GetColumnName(targetCol)}{targetRow}"].Formula = formula;
            Spreadsheet.ActiveGrid.InvalidateCell(targetRow, targetCol);

            MessageBox.Show($"合計を {GetColumnName(targetCol)}{targetRow} に計算しました。", "合計", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"合計の計算に失敗しました。\n{ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void OnAddClick(object sender, RoutedEventArgs e)
    {
        InsertFormulaOperation("+", "足し算");
    }

    private void OnSubtractClick(object sender, RoutedEventArgs e)
    {
        InsertFormulaOperation("-", "引き算");
    }

    private void OnMultiplyClick(object sender, RoutedEventArgs e)
    {
        InsertFormulaOperation("*", "掛け算");
    }

    private void OnDivideClick(object sender, RoutedEventArgs e)
    {
        InsertFormulaOperation("/", "割り算");
    }

    private void OnAverageClick(object sender, RoutedEventArgs e)
    {
        try
        {
            var selectedRanges = Spreadsheet.ActiveGrid.SelectedRanges;
            if (selectedRanges == null || selectedRanges.Count == 0)
            {
                MessageBox.Show("セルを選択してから平均ボタンを押してください。", "平均", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var range = selectedRanges[0];
            var worksheet = Spreadsheet.ActiveSheet;

            // 選択範囲の下または右に結果を出力
            var targetRow = range.Bottom + 1;
            var targetCol = range.Left;

            var rangeAddress = $"{GetColumnName(range.Left)}{range.Top}:{GetColumnName(range.Right)}{range.Bottom}";
            var formula = $"=AVERAGE({rangeAddress})";

            worksheet[$"{GetColumnName(targetCol)}{targetRow}"].Formula = formula;
            Spreadsheet.ActiveGrid.InvalidateCell(targetRow, targetCol);

            MessageBox.Show($"平均を {GetColumnName(targetCol)}{targetRow} に計算しました。", "平均", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"平均の計算に失敗しました。\n{ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void InsertFormulaOperation(string operation, string operationName)
    {
        try
        {
            var selectedRanges = Spreadsheet.ActiveGrid.SelectedRanges;
            if (selectedRanges == null || selectedRanges.Count == 0)
            {
                MessageBox.Show($"2つ以上のセルを選択してから{operationName}ボタンを押してください。", operationName, MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var range = selectedRanges[0];
            if ((range.Right - range.Left + 1) * (range.Bottom - range.Top + 1) < 2)
            {
                MessageBox.Show($"2つ以上のセルを選択してください。", operationName, MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var worksheet = Spreadsheet.ActiveSheet;
            var cells = new List<string>();

            // 選択範囲のセルアドレスを収集
            for (int row = range.Top; row <= range.Bottom; row++)
            {
                for (int col = range.Left; col <= range.Right; col++)
                {
                    cells.Add($"{GetColumnName(col)}{row}");
                }
            }

            // 数式を作成
            var formula = "=" + string.Join(operation, cells);

            // 選択範囲の次のセルに結果を出力
            var targetRow = range.Bottom + 1;
            var targetCol = range.Left;

            worksheet[$"{GetColumnName(targetCol)}{targetRow}"].Formula = formula;
            Spreadsheet.ActiveGrid.InvalidateCell(targetRow, targetCol);

            MessageBox.Show($"{operationName}の結果を {GetColumnName(targetCol)}{targetRow} に計算しました。", operationName, MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"{operationName}の計算に失敗しました。\n{ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private string GetColumnName(int columnIndex)
    {
        // 列番号をアルファベットに変換（1=A, 2=B, ...）
        if (columnIndex <= 0) return "A";
        if (columnIndex <= 26) return ((char)('A' + columnIndex - 1)).ToString();

        // 26以上の場合（AA, AB, ...）
        var result = "";
        while (columnIndex > 0)
        {
            var remainder = (columnIndex - 1) % 26;
            result = (char)('A' + remainder) + result;
            columnIndex = (columnIndex - 1) / 26;
        }
        return result;
    }

    private void OnVoiceInputClick(object sender, RoutedEventArgs e)
    {
        // Windows音声入力を起動（Win+H）
        try
        {
            var speechService = App.Services?.GetService(typeof(ISpeechService)) as ISpeechService;
            if (speechService != null)
            {
                speechService.ActivateWindowsVoiceTyping();
                MessageBox.Show("音声入力を開始しました。\n話し終わったら「停止」と言ってください。", "音声入力", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"音声入力を開始できませんでした。\n{ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void OnHelpClick(object sender, RoutedEventArgs e)
    {
        var helpText = @"【表モードの使い方】

■ 基本操作
　数字を入れて計算ができます。
　家計簿や名簿を作るときに使います。

■ 計算ボタン
　・合計(SUM): セルを選択して押すと合計を計算
　・平均: 選択範囲の平均を計算

■ テンキー
　右側のテンキーで数字や計算式を簡単に入力できます。
　・数字ボタン: 選択中のセルに数字を入力
　・+、-、×、÷: 計算式を作成
　・=: 計算式の最初に付ける
　・C: セルの内容をクリア

■ コマンド入力
　話しかけるだけで操作できます。
　例：「A2に1万円入れて」
　　　「A1からA3を足してA4に」

■ 印刷
　「印刷」ボタンで全てを1ページに収めて印刷できます。";

        MessageBox.Show(helpText, "ヘルプ", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    // ========================================
    // テンキーハンドラー
    // ========================================

    private void OnNumericKeyClick(object sender, RoutedEventArgs e)
    {
        if (sender is Button button && button.Tag is string value)
        {
            try
            {
                var activeCell = Spreadsheet.ActiveGrid.CurrentCell;
                var worksheet = Spreadsheet.ActiveSheet;

                if (activeCell != null)
                {
                    var row = activeCell.RowIndex;
                    var col = activeCell.ColumnIndex;
                    var cellAddress = $"{GetColumnName(col)}{row}";

                    // 現在のセルの値を取得
                    var currentValue = worksheet[cellAddress].Value?.ToString() ?? "";

                    // 数値を追加
                    var newValue = currentValue + value;
                    worksheet[cellAddress].Value = newValue;
                    Spreadsheet.ActiveGrid.InvalidateCell(row, col);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"入力に失敗しました。\n{ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    private void OnOperatorClick(object sender, RoutedEventArgs e)
    {
        if (sender is Button button && button.Tag is string op)
        {
            try
            {
                var activeCell = Spreadsheet.ActiveGrid.CurrentCell;
                var worksheet = Spreadsheet.ActiveSheet;

                if (activeCell != null)
                {
                    var row = activeCell.RowIndex;
                    var col = activeCell.ColumnIndex;
                    var cellAddress = $"{GetColumnName(col)}{row}";

                    // 現在のセルの値を取得
                    var currentValue = worksheet[cellAddress].Value?.ToString() ?? "";

                    // 演算子を追加（計算式として）
                    string newValue;
                    if (string.IsNullOrEmpty(currentValue))
                    {
                        // 空の場合は = から始める
                        newValue = "=";
                    }
                    else if (!currentValue.StartsWith("="))
                    {
                        // 数式でない場合は、=現在値+演算子 にする
                        newValue = $"={currentValue}{op}";
                    }
                    else
                    {
                        // すでに数式の場合は演算子を追加
                        newValue = currentValue + op;
                    }

                    worksheet[cellAddress].Value = newValue;
                    Spreadsheet.ActiveGrid.InvalidateCell(row, col);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"演算子の入力に失敗しました。\n{ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    private void OnEqualsClick(object sender, RoutedEventArgs e)
    {
        try
        {
            var activeCell = Spreadsheet.ActiveGrid.CurrentCell;
            var worksheet = Spreadsheet.ActiveSheet;

            if (activeCell != null)
            {
                var row = activeCell.RowIndex;
                var col = activeCell.ColumnIndex;
                var cellAddress = $"{GetColumnName(col)}{row}";

                // 現在のセルの値を取得
                var currentValue = worksheet[cellAddress].Value?.ToString() ?? "";

                // = を追加（計算式の開始）
                if (!currentValue.StartsWith("="))
                {
                    worksheet[cellAddress].Value = "=" + currentValue;
                    Spreadsheet.ActiveGrid.InvalidateCell(row, col);
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"入力に失敗しました。\n{ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void OnClearClick(object sender, RoutedEventArgs e)
    {
        try
        {
            var activeCell = Spreadsheet.ActiveGrid.CurrentCell;
            var worksheet = Spreadsheet.ActiveSheet;

            if (activeCell != null)
            {
                var row = activeCell.RowIndex;
                var col = activeCell.ColumnIndex;
                var cellAddress = $"{GetColumnName(col)}{row}";

                // セルをクリア
                worksheet[cellAddress].Clear();
                Spreadsheet.ActiveGrid.InvalidateCell(row, col);

                MessageBox.Show($"{cellAddress} をクリアしました。", "クリア", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"クリアに失敗しました。\n{ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }
}
