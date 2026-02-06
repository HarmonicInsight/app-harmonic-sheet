using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using HarmonicSheet.Services;
using Microsoft.Win32;

namespace HarmonicSheet.Views;

public partial class DocumentView : UserControl
{
    private string? _currentFilePath;
    private static readonly string RecentFilesPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "HarmonicSheet",
        "recent_documents.txt");

    public DocumentView()
    {
        InitializeComponent();

        // サービスの取得を試みる
        if (Application.Current is App)
        {
            // DIからサービスを取得する場合はここで設定
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
                var extension = Path.GetExtension(recentFiles[0]).ToLowerInvariant();
                if (extension == ".docx")
                {
                    using var stream = File.OpenRead(recentFiles[0]);
                    RichTextEditor.Load(stream, Syncfusion.Windows.Controls.RichTextBoxAdv.FormatType.Docx);
                }
                else
                {
                    var text = File.ReadAllText(recentFiles[0]);
                    RichTextEditor.Document.Sections.Clear();
                    var section = new Syncfusion.Windows.Controls.RichTextBoxAdv.SectionAdv();
                    var paragraph = new Syncfusion.Windows.Controls.RichTextBoxAdv.ParagraphAdv();
                    var span = new Syncfusion.Windows.Controls.RichTextBoxAdv.SpanAdv { Text = text };
                    paragraph.Inlines.Add(span);
                    section.Blocks.Add(paragraph);
                    RichTextEditor.Document.Sections.Add(section);
                }
                _currentFilePath = recentFiles[0];
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
            "新しい文書を作りますか？\n今の内容は消えます。",
            "確認",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question);

        if (result == MessageBoxResult.Yes)
        {
            RichTextEditor.Document.Sections.Clear();
            _currentFilePath = null;
        }
    }

    private void OnOpenClick(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Word ファイル (*.docx)|*.docx|テキスト ファイル (*.txt)|*.txt|すべてのファイル (*.*)|*.*",
            Title = "開くファイルを選んでください"
        };

        if (dialog.ShowDialog() == true)
        {
            try
            {
                var extension = Path.GetExtension(dialog.FileName).ToLowerInvariant();

                if (extension == ".docx")
                {
                    using var stream = File.OpenRead(dialog.FileName);
                    RichTextEditor.Load(stream, Syncfusion.Windows.Controls.RichTextBoxAdv.FormatType.Docx);
                }
                else
                {
                    var text = File.ReadAllText(dialog.FileName);
                    RichTextEditor.Document.Sections.Clear();
                    // テキストを追加
                    var section = new Syncfusion.Windows.Controls.RichTextBoxAdv.SectionAdv();
                    var paragraph = new Syncfusion.Windows.Controls.RichTextBoxAdv.ParagraphAdv();
                    var span = new Syncfusion.Windows.Controls.RichTextBoxAdv.SpanAdv { Text = text };
                    paragraph.Inlines.Add(span);
                    section.Blocks.Add(paragraph);
                    RichTextEditor.Document.Sections.Add(section);
                }

                _currentFilePath = dialog.FileName;
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
            Filter = "Word ファイル (*.docx)|*.docx|テキスト ファイル (*.txt)|*.txt",
            Title = "保存するファイル名を入力してください",
            FileName = string.IsNullOrEmpty(_currentFilePath)
                ? "文書"
                : Path.GetFileNameWithoutExtension(_currentFilePath)
        };

        if (dialog.ShowDialog() == true)
        {
            try
            {
                var extension = Path.GetExtension(dialog.FileName).ToLowerInvariant();

                using var stream = File.Create(dialog.FileName);
                if (extension == ".docx")
                {
                    RichTextEditor.Save(stream, Syncfusion.Windows.Controls.RichTextBoxAdv.FormatType.Docx);
                }
                else
                {
                    RichTextEditor.Save(stream, Syncfusion.Windows.Controls.RichTextBoxAdv.FormatType.Txt);
                }

                _currentFilePath = dialog.FileName;
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
            // Use PrintDocumentCommand to print
            var printCommand = Syncfusion.Windows.Controls.RichTextBoxAdv.SfRichTextBoxAdv.PrintDocumentCommand;
            if (printCommand.CanExecute(null, RichTextEditor))
            {
                printCommand.Execute(null, RichTextEditor);
            }
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

    private void OnReadAloudClick(object sender, RoutedEventArgs e)
    {
        try
        {
            // ドキュメントからテキストを取得
            var text = GetDocumentText();

            if (string.IsNullOrWhiteSpace(text))
            {
                return;
            }

            // System.Speechを直接使用
            var synthesizer = new System.Speech.Synthesis.SpeechSynthesizer();

            // 日本語音声を探す
            var japaneseVoice = synthesizer.GetInstalledVoices()
                .FirstOrDefault(v => v.VoiceInfo.Culture.Name.StartsWith("ja"));
            if (japaneseVoice != null)
            {
                synthesizer.SelectVoice(japaneseVoice.VoiceInfo.Name);
            }

            synthesizer.Rate = 0; // ゆっくり
            synthesizer.SpeakAsync(text);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"読み上げできませんでした: {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void OnStopReadingClick(object sender, RoutedEventArgs e)
    {
        // 読み上げ停止（簡易実装）
    }

    private string GetDocumentText()
    {
        // RichTextEditorからテキストを抽出
        var sb = new System.Text.StringBuilder();

        foreach (Syncfusion.Windows.Controls.RichTextBoxAdv.SectionAdv section in RichTextEditor.Document.Sections)
        {
            foreach (var block in section.Blocks)
            {
                if (block is Syncfusion.Windows.Controls.RichTextBoxAdv.ParagraphAdv paragraph)
                {
                    foreach (var inline in paragraph.Inlines)
                    {
                        if (inline is Syncfusion.Windows.Controls.RichTextBoxAdv.SpanAdv span)
                        {
                            sb.Append(span.Text);
                        }
                    }
                    sb.AppendLine();
                }
            }
        }

        return sb.ToString();
    }

    private void OnCommandKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter && !Keyboard.Modifiers.HasFlag(ModifierKeys.Shift))
        {
            OnExecuteCommand(sender, e);
            e.Handled = true;
        }
    }

    private async void OnExecuteCommand(object sender, RoutedEventArgs e)
    {
        var command = CommandInput.Text?.Trim();
        if (string.IsNullOrEmpty(command))
        {
            return;
        }

        try
        {
            // ここでClaudeサービスを呼び出す（将来実装）
            // 今は簡易メッセージを表示
            await Task.Delay(500);
            MessageBox.Show("AIアシスタントは設定画面でAPIキーを入力すると使えます", "情報", MessageBoxButton.OK, MessageBoxImage.Information);
            CommandInput.Text = string.Empty;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"エラー: {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
        }
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

    private void OnVoiceInputClick(object sender, RoutedEventArgs e)
    {
        // Windows音声入力を起動（Win+H）
        try
        {
            // まず文書エディタにフォーカスを設定（Windowsの音声入力に必要）
            RichTextEditor.Focus();

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
        var helpText = @"【文書モードの使い方】

■ 基本操作
　文章を書くことができます。
　手紙や報告書を作るときに使います。

■ コマンド入力
　AIに話しかけて文章を編集できます。
　例：「挨拶文を追加して」「敬語に直して」

■ 印刷
　「印刷」ボタンで印刷できます。

■ 音声入力
　マイクボタンを押すと声で文字を入力できます。";

        MessageBox.Show(helpText, "ヘルプ", MessageBoxButton.OK, MessageBoxImage.Information);
    }
}
