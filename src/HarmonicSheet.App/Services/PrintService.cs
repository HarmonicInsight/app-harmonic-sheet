using System.Printing;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using HarmonicSheet.Models;

namespace HarmonicSheet.Services;

/// <summary>
/// 印刷サービスの実装
/// シニア向けに大きなフォントで印刷
/// </summary>
public class PrintService : IPrintService
{
    // シニア向けの大きなフォントサイズ
    private const double PrintFontSize = 14;
    private readonly FontFamily _printFont = new("Yu Gothic UI, Meiryo, MS Gothic");

    public void PrintSpreadsheet(List<SpreadsheetRow> rows)
    {
        var printDialog = new PrintDialog();
        if (printDialog.ShowDialog() != true)
            return;

        var document = CreateSpreadsheetDocument(rows);
        printDialog.PrintDocument(((IDocumentPaginatorSource)document).DocumentPaginator, "表の印刷");
    }

    public void PrintDocument(string text, string title)
    {
        var printDialog = new PrintDialog();
        if (printDialog.ShowDialog() != true)
            return;

        var document = CreateTextDocument(text, title);
        printDialog.PrintDocument(((IDocumentPaginatorSource)document).DocumentPaginator, title);
    }

    public void PrintMail(MailMessage mail)
    {
        var printDialog = new PrintDialog();
        if (printDialog.ShowDialog() != true)
            return;

        var document = CreateMailDocument(mail);
        printDialog.PrintDocument(((IDocumentPaginatorSource)document).DocumentPaginator, "メールの印刷");
    }

    private FlowDocument CreateSpreadsheetDocument(List<SpreadsheetRow> rows)
    {
        var document = new FlowDocument
        {
            FontFamily = _printFont,
            FontSize = PrintFontSize,
            PagePadding = new Thickness(50)
        };

        // タイトル
        var title = new Paragraph(new Run("表データ"))
        {
            FontSize = 20,
            FontWeight = FontWeights.Bold,
            TextAlignment = TextAlignment.Center,
            Margin = new Thickness(0, 0, 0, 20)
        };
        document.Blocks.Add(title);

        // テーブル作成
        var table = new Table
        {
            CellSpacing = 0,
            BorderBrush = Brushes.Black,
            BorderThickness = new Thickness(1)
        };

        if (rows.Any() && rows[0].Cells.Any())
        {
            // 列定義
            foreach (var _ in rows[0].Cells)
            {
                table.Columns.Add(new TableColumn { Width = new GridLength(100) });
            }

            var rowGroup = new TableRowGroup();

            // ヘッダー行（列名）
            var headerRow = new TableRow { Background = Brushes.LightGray };
            foreach (var cell in rows[0].Cells)
            {
                var headerCell = new TableCell(new Paragraph(new Run(cell.Column.ToString()))
                {
                    FontWeight = FontWeights.Bold,
                    TextAlignment = TextAlignment.Center
                })
                {
                    BorderBrush = Brushes.Black,
                    BorderThickness = new Thickness(0.5),
                    Padding = new Thickness(5)
                };
                headerRow.Cells.Add(headerCell);
            }
            rowGroup.Rows.Add(headerRow);

            // データ行
            foreach (var row in rows)
            {
                var tableRow = new TableRow();
                foreach (var cell in row.Cells)
                {
                    var value = !string.IsNullOrEmpty(cell.DisplayValue) ? cell.DisplayValue
                        : !string.IsNullOrEmpty(cell.Value) ? cell.Value
                        : "";

                    var tableCell = new TableCell(new Paragraph(new Run(value))
                    {
                        TextAlignment = TextAlignment.Right
                    })
                    {
                        BorderBrush = Brushes.Black,
                        BorderThickness = new Thickness(0.5),
                        Padding = new Thickness(5)
                    };
                    tableRow.Cells.Add(tableCell);
                }
                rowGroup.Rows.Add(tableRow);
            }

            table.RowGroups.Add(rowGroup);
        }

        document.Blocks.Add(table);

        // 印刷日時
        var footer = new Paragraph(new Run($"印刷日時: {DateTime.Now:yyyy年M月d日 H:mm}"))
        {
            FontSize = 10,
            Foreground = Brushes.Gray,
            TextAlignment = TextAlignment.Right,
            Margin = new Thickness(0, 20, 0, 0)
        };
        document.Blocks.Add(footer);

        return document;
    }

    private FlowDocument CreateTextDocument(string text, string title)
    {
        var document = new FlowDocument
        {
            FontFamily = _printFont,
            FontSize = PrintFontSize,
            PagePadding = new Thickness(50)
        };

        // タイトル
        var titleParagraph = new Paragraph(new Run(title))
        {
            FontSize = 20,
            FontWeight = FontWeights.Bold,
            TextAlignment = TextAlignment.Center,
            Margin = new Thickness(0, 0, 0, 20)
        };
        document.Blocks.Add(titleParagraph);

        // 本文
        var lines = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
        foreach (var line in lines)
        {
            var paragraph = new Paragraph(new Run(line))
            {
                Margin = new Thickness(0, 0, 0, 10)
            };
            document.Blocks.Add(paragraph);
        }

        // 印刷日時
        var footer = new Paragraph(new Run($"印刷日時: {DateTime.Now:yyyy年M月d日 H:mm}"))
        {
            FontSize = 10,
            Foreground = Brushes.Gray,
            TextAlignment = TextAlignment.Right,
            Margin = new Thickness(0, 20, 0, 0)
        };
        document.Blocks.Add(footer);

        return document;
    }

    private FlowDocument CreateMailDocument(MailMessage mail)
    {
        var document = new FlowDocument
        {
            FontFamily = _printFont,
            FontSize = PrintFontSize,
            PagePadding = new Thickness(50)
        };

        // ヘッダー情報
        var header = new Section();
        header.Blocks.Add(new Paragraph(new Run($"送信者: {mail.From}")) { FontWeight = FontWeights.Bold });
        header.Blocks.Add(new Paragraph(new Run($"宛先: {mail.To}")));
        header.Blocks.Add(new Paragraph(new Run($"日時: {mail.Date:yyyy年M月d日 H:mm}")));
        header.Blocks.Add(new Paragraph(new Run($"件名: {mail.Subject}")) { FontWeight = FontWeights.Bold, FontSize = 18 });

        // 区切り線
        header.Blocks.Add(new Paragraph { BorderBrush = Brushes.Gray, BorderThickness = new Thickness(0, 0, 0, 1), Margin = new Thickness(0, 10, 0, 10) });

        document.Blocks.Add(header);

        // 本文
        var lines = mail.Body.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
        foreach (var line in lines)
        {
            var paragraph = new Paragraph(new Run(line))
            {
                Margin = new Thickness(0, 0, 0, 5)
            };
            document.Blocks.Add(paragraph);
        }

        // 印刷日時
        var footer = new Paragraph(new Run($"印刷日時: {DateTime.Now:yyyy年M月d日 H:mm}"))
        {
            FontSize = 10,
            Foreground = Brushes.Gray,
            TextAlignment = TextAlignment.Right,
            Margin = new Thickness(0, 20, 0, 0)
        };
        document.Blocks.Add(footer);

        return document;
    }
}
