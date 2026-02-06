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

    public DocumentView()
    {
        InitializeComponent();

        // サービスの取得を試みる
        if (Application.Current is App)
        {
            // DIからサービスを取得する場合はここで設定
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
            FileNameText.Text = "新しい文書";
            StatusText.Text = "新しい文書を作りました";
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
                FileNameText.Text = Path.GetFileName(dialog.FileName);
                StatusText.Text = $"ファイルを開きました: {Path.GetFileName(dialog.FileName)}";
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
                FileNameText.Text = Path.GetFileName(dialog.FileName);
                StatusText.Text = $"保存しました: {Path.GetFileName(dialog.FileName)}";
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
                StatusText.Text = "印刷しました";
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
                StatusText.Text = "読み上げる文章がありません";
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

            StatusText.Text = "読み上げ中...";
        }
        catch (Exception ex)
        {
            StatusText.Text = $"読み上げできませんでした: {ex.Message}";
        }
    }

    private void OnStopReadingClick(object sender, RoutedEventArgs e)
    {
        // 読み上げ停止（簡易実装）
        StatusText.Text = "読み上げを停止しました";
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

        StatusText.Text = "AIに依頼中...";

        try
        {
            // ここでClaudeサービスを呼び出す（将来実装）
            // 今は簡易メッセージを表示
            await Task.Delay(500);
            StatusText.Text = "AIアシスタントは設定画面でAPIキーを入力すると使えます";
            CommandInput.Text = string.Empty;
        }
        catch (Exception ex)
        {
            StatusText.Text = $"エラー: {ex.Message}";
        }
    }
}
