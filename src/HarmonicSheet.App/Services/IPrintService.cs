using HarmonicSheet.Models;

namespace HarmonicSheet.Services;

/// <summary>
/// 印刷サービスのインターフェース
/// </summary>
public interface IPrintService
{
    /// <summary>
    /// スプレッドシートを印刷します
    /// </summary>
    void PrintSpreadsheet(List<SpreadsheetRow> rows);

    /// <summary>
    /// 文書を印刷します
    /// </summary>
    void PrintDocument(string text, string title);

    /// <summary>
    /// メールを印刷します
    /// </summary>
    void PrintMail(MailMessage mail);
}
