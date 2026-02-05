using HarmonicSheet.Models;

namespace HarmonicSheet.Services;

/// <summary>
/// スプレッドシートファイル操作サービスのインターフェース
/// </summary>
public interface ISpreadsheetService
{
    /// <summary>
    /// Excelファイルに保存します
    /// </summary>
    Task SaveToExcel(List<SpreadsheetRow> rows, string filePath);

    /// <summary>
    /// Excelファイルから読み込みます
    /// </summary>
    Task<List<SpreadsheetRow>> LoadFromExcel(string filePath);
}
