namespace HarmonicSheet.Services;

/// <summary>
/// 文書ファイル操作サービスのインターフェース
/// </summary>
public interface IDocumentService
{
    /// <summary>
    /// 文書を保存します
    /// </summary>
    Task SaveDocument(string text, string filePath);

    /// <summary>
    /// 文書を読み込みます
    /// </summary>
    Task<string> LoadDocument(string filePath);
}
