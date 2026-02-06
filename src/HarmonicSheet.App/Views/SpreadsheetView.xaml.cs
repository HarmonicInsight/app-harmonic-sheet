using System.Collections.ObjectModel;
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

    // キーパッドモード
    private enum KeypadMode
    {
        Calculator,  // 計算機モード
        Input        // 入力モード（セルに式を入力）
    }
    private KeypadMode _keypadMode = KeypadMode.Calculator;

    // 電卓の状態
    private string _calcCurrentValue = "0";
    private string _calcOperator = "";
    private double _calcStoredValue = 0;
    private bool _calcNewNumber = true;

    // 計算履歴
    private ObservableCollection<string> _calcHistoryList = new ObservableCollection<string>();

    // 入力モードの状態
    private string _formulaBuffer = "";

    public SpreadsheetView()
    {
        InitializeComponent();

        // DIが使える場合はサービスを取得
        if (Application.Current is App app)
        {
            // サービスは後で設定される
        }

        // 計算履歴のバインディング
        CalcHistory.ItemsSource = _calcHistoryList;

        // 新規ワークブックを作成（1シート）
        try
        {
            Spreadsheet.Create(1);

            // シニア向けの大きなフォント設定
            ConfigureForSeniors();

            // モードUIを初期化
            UpdateModeUI();
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

                // 自動計算を有効化
                worksheet.EnableSheetCalculations();
            }
        }
        catch
        {
            // 設定に失敗しても続行
        }
    }

    private void OnUndoClick(object sender, RoutedEventArgs e)
    {
        try
        {
            // Syncfusion Spreadsheet does not have built-in Undo
            // Implement simple undo using history tracking
            MessageBox.Show(
                "元に戻す機能:\n\n" +
                "・間違えて入力した内容を消してください\n" +
                "・または、前回保存したファイルを開き直してください\n\n" +
                "※ 今後のバージョンで自動保存機能を追加予定です",
                "元に戻す",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"操作に失敗しました。\n{ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
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
            "新しい表を作りますか？\n\n" +
            "・はい → 家計簿テンプレートを使う\n" +
            "・いいえ → まっさらな表を作る\n" +
            "・キャンセル → 何もしない",
            "新しい表",
            MessageBoxButton.YesNoCancel,
            MessageBoxImage.Question);

        if (result == MessageBoxResult.Yes)
        {
            CreateBudgetTemplate();
        }
        else if (result == MessageBoxResult.No)
        {
            Spreadsheet.Create(1);
            ConfigureForSeniors();
        }
    }

    private void CreateBudgetTemplate()
    {
        try
        {
            Spreadsheet.Create(1);
            ConfigureForSeniors();

            var worksheet = Spreadsheet.ActiveSheet;

            // タイトル
            worksheet["A1"].Value = "家計簿";
            worksheet["A1"].CellStyle.Font.Size = 20;
            worksheet["A1"].CellStyle.Font.Bold = true;

            // 月の入力
            worksheet["B1"].Value = DateTime.Now.ToString("yyyy年MM月");
            worksheet["B1"].CellStyle.Font.Size = 16;

            // ヘッダー
            worksheet["A3"].Value = "項目";
            worksheet["B3"].Value = "予算";
            worksheet["C3"].Value = "実際";
            worksheet["D3"].Value = "差額";

            // ヘッダーのスタイル
            for (int col = 1; col <= 4; col++)
            {
                var cell = worksheet[$"{GetColumnName(col)}3"];
                cell.CellStyle.Font.Bold = true;
                cell.CellStyle.Font.Size = 14;
                cell.CellStyle.ColorIndex = Syncfusion.XlsIO.ExcelKnownColors.Grey_25_percent;
            }

            // 項目
            var items = new[] { "食費", "光熱費", "水道代", "電気代", "ガス代", "通信費", "医療費", "交通費", "その他" };
            int row = 4;
            foreach (var item in items)
            {
                worksheet[$"A{row}"].Value = item;
                worksheet[$"D{row}"].Formula = $"=B{row}-C{row}"; // 差額計算
                row++;
            }

            // 合計行
            worksheet[$"A{row}"].Value = "合計";
            worksheet[$"A{row}"].CellStyle.Font.Bold = true;
            worksheet[$"B{row}"].Formula = $"=SUM(B4:B{row - 1})";
            worksheet[$"C{row}"].Formula = $"=SUM(C4:C{row - 1})";
            worksheet[$"D{row}"].Formula = $"=SUM(D4:D{row - 1})";

            // 再計算を強制
            worksheet.Calculate();

            // 列幅調整（既にConfigureForSeniorsで設定済み）

            MessageBox.Show(
                "家計簿テンプレートを作りました！\n\n" +
                "【使い方】\n" +
                "1. B列に「予算」を入力\n" +
                "2. C列に「実際の金額」を入力\n" +
                "3. D列に自動で「差額」が計算されます\n\n" +
                "プラスなら節約、マイナスなら予算オーバーです",
                "家計簿テンプレート",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"テンプレートの作成に失敗しました。\n{ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
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

            // 再計算を強制
            worksheet.Calculate();
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

            // 再計算を強制
            worksheet.Calculate();
            Spreadsheet.ActiveGrid.InvalidateCell(targetRow, targetCol);

            MessageBox.Show($"平均を {GetColumnName(targetCol)}{targetRow} に計算しました。", "平均", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"平均の計算に失敗しました。\n{ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void OnTax10Click(object sender, RoutedEventArgs e)
    {
        ApplyTaxCalculation(1.10, "10%の消費税");
    }

    private void OnTax8Click(object sender, RoutedEventArgs e)
    {
        ApplyTaxCalculation(1.08, "8%の消費税");
    }

    private void OnDiscountClick(object sender, RoutedEventArgs e)
    {
        try
        {
            var input = Microsoft.VisualBasic.Interaction.InputBox(
                "何%割引ですか？\n（例：30 と入力すると30%オフ）",
                "割引計算",
                "30");

            if (string.IsNullOrWhiteSpace(input))
                return;

            if (double.TryParse(input, out var discountPercent) && discountPercent > 0 && discountPercent < 100)
            {
                var multiplier = (100 - discountPercent) / 100;
                ApplyTaxCalculation(multiplier, $"{discountPercent}%割引");
            }
            else
            {
                MessageBox.Show("0から100の間の数字を入れてください。", "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"割引計算に失敗しました。\n{ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void OnCumulativeClick(object sender, RoutedEventArgs e)
    {
        try
        {
            var selectedRanges = Spreadsheet.ActiveGrid.SelectedRanges;
            if (selectedRanges == null || selectedRanges.Count == 0)
            {
                MessageBox.Show("セルを選択してから累計ボタンを押してください。", "累計計算", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var range = selectedRanges[0];
            var worksheet = Spreadsheet.ActiveSheet;

            // 範囲の各セルに累計を表示
            if (range.Top == range.Bottom)
            {
                // 横一列の場合
                double cumulative = 0;
                var targetRow = range.Bottom + 1;

                for (int col = range.Left; col <= range.Right; col++)
                {
                    var cellAddress = $"{GetColumnName(col)}{range.Top}";
                    var cell = worksheet[cellAddress];

                    if (cell.Value != null && double.TryParse(cell.Value.ToString(), out var value))
                    {
                        cumulative += value;
                        var targetAddress = $"{GetColumnName(col)}{targetRow}";
                        worksheet[targetAddress].Value = cumulative.ToString();
                        Spreadsheet.ActiveGrid.InvalidateCell(targetRow, col);
                    }
                }

                MessageBox.Show($"累計を計算しました。\n下の行に累計が表示されています。", "累計計算完了", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else if (range.Left == range.Right)
            {
                // 縦一列の場合
                double cumulative = 0;
                var targetCol = range.Right + 1;

                for (int row = range.Top; row <= range.Bottom; row++)
                {
                    var cellAddress = $"{GetColumnName(range.Left)}{row}";
                    var cell = worksheet[cellAddress];

                    if (cell.Value != null && double.TryParse(cell.Value.ToString(), out var value))
                    {
                        cumulative += value;
                        var targetAddress = $"{GetColumnName(targetCol)}{row}";
                        worksheet[targetAddress].Value = cumulative.ToString();
                        Spreadsheet.ActiveGrid.InvalidateCell(row, targetCol);
                    }
                }

                MessageBox.Show($"累計を計算しました。\n右の列に累計が表示されています。", "累計計算完了", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                MessageBox.Show("一行または一列を選んでください。", "累計計算", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"累計計算に失敗しました。\n{ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void ApplyTaxCalculation(double multiplier, string operation)
    {
        try
        {
            var selectedRanges = Spreadsheet.ActiveGrid.SelectedRanges;
            if (selectedRanges == null || selectedRanges.Count == 0)
            {
                MessageBox.Show("セルを選択してからボタンを押してください。", "消費税計算", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var range = selectedRanges[0];
            var worksheet = Spreadsheet.ActiveSheet;

            // 単一セルの場合
            if (range.Left == range.Right && range.Top == range.Bottom)
            {
                var row = range.Top;
                var col = range.Left;
                var cellAddress = $"{GetColumnName(col)}{row}";
                var cell = worksheet[cellAddress];

                // 現在の値を取得して確認
                if (cell.Value != null && double.TryParse(cell.Value.ToString(), out var value))
                {
                    var targetRow = row + 1;
                    var targetAddress = $"{GetColumnName(col)}{targetRow}";

                    // 計算式を挿入（直接計算せず、式を入れる）
                    var formula = $"={cellAddress}*{multiplier}";
                    worksheet[targetAddress].Formula = formula;

                    // 再計算を強制
                    worksheet.Calculate();
                    Spreadsheet.ActiveGrid.InvalidateCell(targetRow, col);

                    MessageBox.Show($"{operation}の計算式を {targetAddress} に入力しました。\n式: {formula}",
                        "計算完了", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    MessageBox.Show("数字が入っているセルを選んでください。", "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            else
            {
                // 範囲選択の場合
                var sourceAddress = $"{GetColumnName(range.Left)}{range.Top}:{GetColumnName(range.Right)}{range.Bottom}";
                var targetRow = range.Bottom + 1;
                var targetCol = range.Left;
                var targetAddress = $"{GetColumnName(targetCol)}{targetRow}";

                var formula = $"=SUM({sourceAddress})*{multiplier}";
                worksheet[targetAddress].Formula = formula;

                // 再計算を強制
                worksheet.Calculate();
                Spreadsheet.ActiveGrid.InvalidateCell(targetRow, targetCol);

                MessageBox.Show($"{operation}を適用して {targetAddress} に計算しました。",
                    "計算完了", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"計算に失敗しました。\n{ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
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
            // まずテキストボックスにフォーカスを設定（Windowsの音声入力に必要）
            CommandInput.Focus();

            // 少し待ってから音声入力を起動
            Task.Delay(100).ContinueWith(_ =>
            {
                Dispatcher.Invoke(() =>
                {
                    var speechService = App.Services?.GetService(typeof(ISpeechService)) as ISpeechService;
                    if (speechService != null)
                    {
                        speechService.ActivateWindowsVoiceTyping();
                        MessageBox.Show("音声入力を開始しました。\n話し終わったら「停止」と言ってください。", "音声入力", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                });
            });
        }
        catch (Exception ex)
        {
            MessageBox.Show($"音声入力を開始できませんでした。\n{ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void OnHelpClick(object sender, RoutedEventArgs e)
    {
        var helpText = @"🤖 AIアシスタント - 困ったときはこちら

━━━━━━━━━━━━━━━━━━━━━━

💬 よくある質問

Q: 合計を計算したい
A: セルを選んで「合計」ボタンを押してください。
   選んだセルの下に合計が表示されます。

Q: 消費税を計算したい
A: 金額のセルを選んで「消費税10%」か
   「消費税8%」ボタンを押してください。
   計算式が自動で入ります。

Q: 計算式が0になってしまう
A: 計算式を入れた後、少し待ってください。
   自動で再計算されます。

Q: 家計簿を作りたい
A: 「新規」ボタンを押すと、家計簿テンプレート
   が使えます。

━━━━━━━━━━━━━━━━━━━━━━

📞 電話サポート
　0120-XXX-XXX (平日 9:00〜18:00)

💡 使い方のコツ
　・数字を入れたいセルをクリック
　・ボタンを押す前に必ずセルを選択
　・テンキーボタンで数字を簡単入力
　話しかけるだけで操作できます。
　例：「A2に1万円入れて」
　　　「A1からA3を足してA4に」

■ 印刷
　「印刷」ボタンで全てを1ページに収めて印刷できます。";

        MessageBox.Show(helpText, "ヘルプ", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private void OnToggleKeypadClick(object sender, RoutedEventArgs e)
    {
        // テンキーパネルの表示/非表示を切り替え
        if (KeypadPanel.Visibility == Visibility.Visible)
        {
            KeypadPanel.Visibility = Visibility.Collapsed;
            BtnToggleKeypad.Content = new TextBlock { Text = "テンキー表示", FontSize = 24, FontWeight = FontWeights.Bold };
        }
        else
        {
            KeypadPanel.Visibility = Visibility.Visible;
            BtnToggleKeypad.Content = new TextBlock { Text = "テンキー非表示", FontSize = 24, FontWeight = FontWeights.Bold };
        }
    }

    // ========================================
    // モード切り替え
    // ========================================

    private void OnCalcModeClick(object sender, RoutedEventArgs e)
    {
        _keypadMode = KeypadMode.Calculator;
        UpdateModeUI();
    }

    private void OnInputModeClick(object sender, RoutedEventArgs e)
    {
        _keypadMode = KeypadMode.Input;
        UpdateModeUI();
    }

    private void UpdateModeUI()
    {
        if (_keypadMode == KeypadMode.Calculator)
        {
            // 計算機モード
            BtnCalcMode.Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(184, 148, 47)); // Gold
            BtnCalcMode.Foreground = System.Windows.Media.Brushes.White;
            BtnInputMode.Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(231, 226, 218)); // Light gray
            BtnInputMode.Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(28, 25, 23));

            CalcModePanel.Visibility = Visibility.Visible;
            InputModePanel.Visibility = Visibility.Collapsed;
            CalcModeButtons.Visibility = Visibility.Visible;
            InputModeButtons.Visibility = Visibility.Collapsed;
        }
        else
        {
            // 入力モード
            BtnInputMode.Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(184, 148, 47)); // Gold
            BtnInputMode.Foreground = System.Windows.Media.Brushes.White;
            BtnCalcMode.Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(231, 226, 218)); // Light gray
            BtnCalcMode.Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(28, 25, 23));

            CalcModePanel.Visibility = Visibility.Collapsed;
            InputModePanel.Visibility = Visibility.Visible;
            CalcModeButtons.Visibility = Visibility.Collapsed;
            InputModeButtons.Visibility = Visibility.Visible;
        }
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
                if (_keypadMode == KeypadMode.Calculator)
                {
                    // 計算機モードで数字を入力
                    if (_calcNewNumber)
                    {
                        _calcCurrentValue = value;
                        _calcNewNumber = false;
                    }
                    else
                    {
                        _calcCurrentValue += value;
                    }
                    UpdateCalcDisplay();
                }
                else
                {
                    // 入力モードで数字を式に追加
                    _formulaBuffer += value;
                    UpdateFormulaDisplay();
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
                if (_keypadMode == KeypadMode.Calculator)
                {
                    // 計算機モード: 前の演算を実行
                    if (!string.IsNullOrEmpty(_calcOperator))
                    {
                        PerformCalculation();
                    }
                    else
                    {
                        _calcStoredValue = double.Parse(_calcCurrentValue);
                    }

                    _calcOperator = op;
                    _calcNewNumber = true;
                    UpdateCalcDisplay();
                }
                else
                {
                    // 入力モード: 演算子を式に追加
                    _formulaBuffer += op;
                    UpdateFormulaDisplay();
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
            if (_keypadMode == KeypadMode.Calculator)
            {
                // 電卓の計算を実行
                if (!string.IsNullOrEmpty(_calcOperator))
                {
                    PerformCalculation();
                    _calcOperator = "";
                    UpdateCalcDisplay();
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"計算に失敗しました。\n{ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void OnClearClick(object sender, RoutedEventArgs e)
    {
        try
        {
            if (_keypadMode == KeypadMode.Calculator)
            {
                // 計算機をクリア
                _calcCurrentValue = "0";
                _calcOperator = "";
                _calcStoredValue = 0;
                _calcNewNumber = true;
                UpdateCalcDisplay();
            }
            else
            {
                // 入力モードをクリア
                _formulaBuffer = "";
                CellSelectIndicator.Visibility = Visibility.Collapsed;
                UpdateFormulaDisplay();
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"クリアに失敗しました。\n{ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    // ========================================
    // 入力モード専用ハンドラー
    // ========================================

    private void OnStartFormulaClick(object sender, RoutedEventArgs e)
    {
        if (_keypadMode == KeypadMode.Input)
        {
            if (string.IsNullOrEmpty(_formulaBuffer) || !_formulaBuffer.StartsWith("="))
            {
                _formulaBuffer = "=" + _formulaBuffer;
                UpdateFormulaDisplay();
            }
        }
    }

    private void OnCellSelectClick(object sender, RoutedEventArgs e)
    {
        try
        {
            // 現在選択されているセルを取得して式に追加
            var selectedRanges = Spreadsheet.ActiveGrid.SelectedRanges;
            if (selectedRanges != null && selectedRanges.Count > 0)
            {
                var range = selectedRanges[0];

                // 単一セルの場合
                if (range.Left == range.Right && range.Top == range.Bottom)
                {
                    var cellAddress = $"{GetColumnName(range.Left)}{range.Top}";
                    _formulaBuffer += cellAddress;
                }
                else
                {
                    // 範囲の場合
                    var rangeAddress = $"{GetColumnName(range.Left)}{range.Top}:{GetColumnName(range.Right)}{range.Bottom}";
                    _formulaBuffer += rangeAddress;
                }

                UpdateFormulaDisplay();

                // インジケーターを一時的に表示
                CellSelectIndicator.Visibility = Visibility.Visible;
                Task.Delay(1000).ContinueWith(_ =>
                {
                    Dispatcher.Invoke(() => CellSelectIndicator.Visibility = Visibility.Collapsed);
                });
            }
            else
            {
                MessageBox.Show("先にスプレッドシートのセルを選択してから、このボタンを押してください。", "セル選択", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"セル参照の追加に失敗しました。\n{ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void OnConfirmFormulaClick(object sender, RoutedEventArgs e)
    {
        try
        {
            if (string.IsNullOrEmpty(_formulaBuffer))
            {
                MessageBox.Show("計算式が空です。", "確認", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var activeCell = Spreadsheet.ActiveGrid.CurrentCell;
            if (activeCell != null)
            {
                var worksheet = Spreadsheet.ActiveSheet;
                var row = activeCell.RowIndex;
                var col = activeCell.ColumnIndex;
                var cellAddress = $"{GetColumnName(col)}{row}";

                // 式をセルに入力
                if (_formulaBuffer.StartsWith("="))
                {
                    worksheet[cellAddress].Formula = _formulaBuffer;
                }
                else
                {
                    worksheet[cellAddress].Value = _formulaBuffer;
                }

                Spreadsheet.ActiveGrid.InvalidateCell(row, col);

                MessageBox.Show($"計算式 {_formulaBuffer} を {cellAddress} に入力しました。", "完了", MessageBoxButton.OK, MessageBoxImage.Information);

                // クリア
                _formulaBuffer = "";
                UpdateFormulaDisplay();
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"計算式の入力に失敗しました。\n{ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void OnCopyResultToCell(object sender, RoutedEventArgs e)
    {
        try
        {
            var activeCell = Spreadsheet.ActiveGrid.CurrentCell;
            if (activeCell != null)
            {
                var worksheet = Spreadsheet.ActiveSheet;
                var row = activeCell.RowIndex;
                var col = activeCell.ColumnIndex;
                var cellAddress = $"{GetColumnName(col)}{row}";

                worksheet[cellAddress].Value = _calcCurrentValue;
                Spreadsheet.ActiveGrid.InvalidateCell(row, col);

                MessageBox.Show($"計算結果 {_calcCurrentValue} を {cellAddress} にコピーしました。", "完了", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"セルへのコピーに失敗しました。\n{ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void OnBackspaceClick(object sender, RoutedEventArgs e)
    {
        try
        {
            if (_formulaBuffer.Length > 0)
            {
                _formulaBuffer = _formulaBuffer.Substring(0, _formulaBuffer.Length - 1);
                UpdateFormulaDisplay();
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"削除に失敗しました。\n{ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void OnFormulaDisplayTextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
    {
        // FormulaDisplayの直接編集を_formulaBufferに反映
        if (sender is System.Windows.Controls.TextBox textBox)
        {
            _formulaBuffer = textBox.Text;
        }
    }

    // ========================================
    // 電卓機能
    // ========================================

    private void UpdateCalcDisplay()
    {
        var displayValue = _calcCurrentValue;
        if (!string.IsNullOrEmpty(_calcOperator) && !_calcNewNumber)
        {
            displayValue = $"{_calcStoredValue} {_calcOperator} {_calcCurrentValue}";
        }
        CalcDisplay.Text = displayValue;
    }

    private void UpdateFormulaDisplay()
    {
        FormulaDisplay.Text = _formulaBuffer;
    }

    private void PerformCalculation()
    {
        var currentNum = double.Parse(_calcCurrentValue);
        double result = _calcOperator switch
        {
            "+" => _calcStoredValue + currentNum,
            "-" => _calcStoredValue - currentNum,
            "*" => _calcStoredValue * currentNum,
            "/" => currentNum != 0 ? _calcStoredValue / currentNum : double.NaN,
            _ => currentNum
        };

        // 計算履歴に追加
        var operatorSymbol = _calcOperator switch
        {
            "+" => "+",
            "-" => "-",
            "*" => "×",
            "/" => "÷",
            _ => ""
        };

        if (!string.IsNullOrEmpty(operatorSymbol))
        {
            var historyEntry = $"{_calcStoredValue} {operatorSymbol} {currentNum} = {result:G}";
            _calcHistoryList.Insert(0, historyEntry); // 最新を先頭に追加

            // 履歴が20件を超えたら古いものを削除
            while (_calcHistoryList.Count > 20)
            {
                _calcHistoryList.RemoveAt(_calcHistoryList.Count - 1);
            }
        }

        _calcCurrentValue = result.ToString("G");
        _calcStoredValue = result;
        _calcNewNumber = true;
    }

    private void OnClearHistoryClick(object sender, RoutedEventArgs e)
    {
        _calcHistoryList.Clear();
        MessageBox.Show("計算履歴をクリアしました。", "履歴クリア", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private void OnHistoryItemDoubleClick(object sender, MouseButtonEventArgs e)
    {
        try
        {
            if (CalcHistory.SelectedItem is string historyEntry)
            {
                // "数値1 演算子 数値2 = 結果" の形式から結果を抽出
                var parts = historyEntry.Split('=');
                if (parts.Length == 2)
                {
                    var resultStr = parts[1].Trim();
                    _calcCurrentValue = resultStr;
                    _calcStoredValue = double.Parse(resultStr);
                    _calcNewNumber = true;
                    _calcOperator = "";
                    UpdateCalcDisplay();
                    MessageBox.Show($"履歴から値 {resultStr} を読み込みました。", "履歴から復元", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"履歴の読み込みに失敗しました。\n{ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

}
