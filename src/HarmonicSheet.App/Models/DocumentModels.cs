namespace HarmonicSheet.Models;

/// <summary>
/// Claude APIからの文書操作結果
/// </summary>
public class DocumentCommandResult
{
    public bool Success { get; set; }
    public string Message { get; set; } = string.Empty;
    public string? NewText { get; set; }
}
