using System.IO;
using System.Text;

namespace HarmonicSheet.Services;

/// <summary>
/// 文書ファイル操作サービスの実装
/// </summary>
public class DocumentService : IDocumentService
{
    public async Task SaveDocument(string text, string filePath)
    {
        // UTF-8（BOM付き）で保存（日本語対応）
        await File.WriteAllTextAsync(filePath, text, new UTF8Encoding(true));
    }

    public async Task<string> LoadDocument(string filePath)
    {
        if (!File.Exists(filePath))
        {
            throw new FileNotFoundException("ファイルが見つかりません", filePath);
        }

        return await File.ReadAllTextAsync(filePath);
    }
}
