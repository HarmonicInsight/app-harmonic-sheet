using HarmonicSheet.Models;

namespace HarmonicSheet.Services;

/// <summary>
/// Claude API（Haiku）連携サービスのインターフェース
/// </summary>
public interface IClaudeService
{
    /// <summary>
    /// APIキーを設定します
    /// </summary>
    void SetApiKey(string apiKey);

    /// <summary>
    /// APIが設定済みかどうか
    /// </summary>
    bool IsConfigured { get; }

    /// <summary>
    /// スプレッドシートの自然言語コマンドを処理します
    /// </summary>
    /// <param name="command">ユーザーの自然言語コマンド（例: "A2に1万円入れて"）</param>
    /// <param name="currentState">現在のグリッドの状態</param>
    /// <returns>操作結果</returns>
    Task<SpreadsheetCommandResult> ProcessSpreadsheetCommand(string command, string currentState);

    /// <summary>
    /// 文書の自然言語コマンドを処理します
    /// </summary>
    /// <param name="command">ユーザーの自然言語コマンド（例: "挨拶文を追加して"）</param>
    /// <param name="currentText">現在の文書テキスト</param>
    /// <returns>操作結果</returns>
    Task<DocumentCommandResult> ProcessDocumentCommand(string command, string currentText);

    /// <summary>
    /// メールの自然言語コマンドを処理します
    /// </summary>
    /// <param name="command">ユーザーの自然言語コマンド（例: "田中さんにメールを書いて"）</param>
    /// <param name="selectedMail">選択中のメール（返信時など）</param>
    /// <returns>操作結果</returns>
    Task<MailCommandResult> ProcessMailCommand(string command, MailMessage? selectedMail);
}
