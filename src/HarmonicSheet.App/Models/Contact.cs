namespace HarmonicSheet.Models;

/// <summary>
/// 連絡先（よく使うメールアドレス）
/// </summary>
public class Contact
{
    public string Id { get; set; } = Guid.NewGuid().ToString();

    /// <summary>表示名（例：田中太郎さん）</summary>
    public string DisplayName { get; set; } = string.Empty;

    /// <summary>メールアドレス</summary>
    public string Email { get; set; } = string.Empty;

    /// <summary>電話番号（任意）</summary>
    public string? PhoneNumber { get; set; }

    /// <summary>メモ（任意）</summary>
    public string? Notes { get; set; }

    /// <summary>グループ（家族、友人、病院など）</summary>
    public string Group { get; set; } = "その他";

    /// <summary>よく使う順（お気に入り）</summary>
    public bool IsFavorite { get; set; }

    /// <summary>最終使用日時</summary>
    public DateTime LastUsed { get; set; } = DateTime.MinValue;

    /// <summary>作成日時</summary>
    public DateTime CreatedAt { get; set; } = DateTime.Now;
}

/// <summary>
/// 連絡先のグループ定義
/// </summary>
public static class ContactGroups
{
    public static readonly string[] DefaultGroups = new[]
    {
        "家族",
        "友人",
        "病院",
        "お店",
        "役所",
        "その他"
    };
}
