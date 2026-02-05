using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using HarmonicSheet.Models;

namespace HarmonicSheet.Services;

/// <summary>
/// Claude API（Haiku）連携サービスの実装
/// シニア向けの自然言語操作をサポート
/// </summary>
public class ClaudeService : IClaudeService, IDisposable
{
    private readonly HttpClient _httpClient;
    private string? _apiKey;
    private bool _disposed;

    // Claude Haiku を使用（高速・低コスト）
    private const string Model = "claude-3-haiku-20240307";
    private const string ApiUrl = "https://api.anthropic.com/v1/messages";

    public ClaudeService()
    {
        _httpClient = new HttpClient();
        _httpClient.DefaultRequestHeaders.Add("anthropic-version", "2023-06-01");
    }

    public bool IsConfigured => !string.IsNullOrEmpty(_apiKey);

    public void SetApiKey(string apiKey)
    {
        _apiKey = apiKey;
        _httpClient.DefaultRequestHeaders.Remove("x-api-key");
        _httpClient.DefaultRequestHeaders.Add("x-api-key", apiKey);
    }

    public async Task<SpreadsheetCommandResult> ProcessSpreadsheetCommand(string command, string currentState)
    {
        if (!IsConfigured)
        {
            return new SpreadsheetCommandResult
            {
                Success = false,
                Message = "APIキーが設定されていません。設定画面でAPIキーを入力してください。"
            };
        }

        var systemPrompt = @"あなたは高齢者向けのスプレッドシート操作アシスタントです。
ユーザーの自然言語での指示を理解し、セルへの値入力や数式の設定を行います。

【重要なルール】
- 「Aの2」「A2」「A列の2行目」は全て同じセル「A2」を指します
- 「1万円」「10000円」「10,000」は全て数値「10000」として扱います
- 「足す」「足して」「合計」「プラス」は全て加算を意味します
- 数式は「=」で始まる形式で返してください（例: =A1+A2+A3）

【応答形式】
以下のJSON形式で応答してください:
{
  ""success"": true,
  ""message"": ""（ユーザーへのわかりやすい説明）"",
  ""changes"": [
    { ""column"": ""A"", ""row"": 2, ""value"": ""10000"", ""formula"": null },
    { ""column"": ""A"", ""row"": 4, ""value"": """", ""formula"": ""=A1+A2+A3"" }
  ]
}

数式を入れる場合はvalueは空文字にしてください。";

        var userMessage = $"現在の表の状態:\n{currentState}\n\n指示: {command}";

        try
        {
            var response = await SendMessage(systemPrompt, userMessage);
            return ParseSpreadsheetResponse(response);
        }
        catch (Exception ex)
        {
            return new SpreadsheetCommandResult
            {
                Success = false,
                Message = $"エラーが発生しました: {ex.Message}"
            };
        }
    }

    public async Task<DocumentCommandResult> ProcessDocumentCommand(string command, string currentText)
    {
        if (!IsConfigured)
        {
            return new DocumentCommandResult
            {
                Success = false,
                Message = "APIキーが設定されていません。設定画面でAPIキーを入力してください。"
            };
        }

        var systemPrompt = @"あなたは高齢者向けの文書作成アシスタントです。
ユーザーの自然言語での指示を理解し、文書の作成・編集を手助けします。

【できること】
- 文章の追加・修正
- 挨拶文や結びの言葉の提案
- 文章の校正
- 文章をわかりやすくする

【応答形式】
以下のJSON形式で応答してください:
{
  ""success"": true,
  ""message"": ""（ユーザーへのわかりやすい説明）"",
  ""newText"": ""（修正後の文書全体、変更がない場合はnull）""
}";

        var userMessage = $"現在の文書:\n{currentText}\n\n指示: {command}";

        try
        {
            var response = await SendMessage(systemPrompt, userMessage);
            return ParseDocumentResponse(response);
        }
        catch (Exception ex)
        {
            return new DocumentCommandResult
            {
                Success = false,
                Message = $"エラーが発生しました: {ex.Message}"
            };
        }
    }

    public async Task<MailCommandResult> ProcessMailCommand(string command, MailMessage? selectedMail)
    {
        if (!IsConfigured)
        {
            return new MailCommandResult
            {
                Success = false,
                Message = "APIキーが設定されていません。設定画面でAPIキーを入力してください。"
            };
        }

        var systemPrompt = @"あなたは高齢者向けのメール作成アシスタントです。
ユーザーの自然言語での指示を理解し、メールの作成を手助けします。

【できること】
- 新規メールの作成
- 返信メールの下書き
- メールの文面の提案

【応答形式】
以下のJSON形式で応答してください:
{
  ""success"": true,
  ""message"": ""（ユーザーへのわかりやすい説明）"",
  ""action"": ""compose"",
  ""to"": ""（宛先メールアドレス、わからない場合は空）"",
  ""subject"": ""（件名）"",
  ""body"": ""（本文）""
}

actionは以下のいずれか:
- compose: 新規メール作成
- reply: 返信
- forward: 転送";

        var context = selectedMail != null
            ? $"選択中のメール:\n送信者: {selectedMail.From}\n件名: {selectedMail.Subject}\n本文: {selectedMail.Body}\n\n"
            : "";

        var userMessage = $"{context}指示: {command}";

        try
        {
            var response = await SendMessage(systemPrompt, userMessage);
            return ParseMailResponse(response);
        }
        catch (Exception ex)
        {
            return new MailCommandResult
            {
                Success = false,
                Message = $"エラーが発生しました: {ex.Message}"
            };
        }
    }

    private async Task<string> SendMessage(string systemPrompt, string userMessage)
    {
        var requestBody = new
        {
            model = Model,
            max_tokens = 1024,
            system = systemPrompt,
            messages = new[]
            {
                new { role = "user", content = userMessage }
            }
        };

        var json = JsonSerializer.Serialize(requestBody);
        var content = new StringContent(json, Encoding.UTF8, "application/json");

        var response = await _httpClient.PostAsync(ApiUrl, content);
        response.EnsureSuccessStatusCode();

        var responseJson = await response.Content.ReadAsStringAsync();
        using var doc = JsonDocument.Parse(responseJson);

        var textContent = doc.RootElement
            .GetProperty("content")[0]
            .GetProperty("text")
            .GetString();

        return textContent ?? string.Empty;
    }

    private SpreadsheetCommandResult ParseSpreadsheetResponse(string response)
    {
        try
        {
            // JSONブロックを抽出（```json ... ``` または { ... }）
            var jsonText = ExtractJson(response);
            using var doc = JsonDocument.Parse(jsonText);
            var root = doc.RootElement;

            var result = new SpreadsheetCommandResult
            {
                Success = root.GetProperty("success").GetBoolean(),
                Message = root.GetProperty("message").GetString() ?? string.Empty
            };

            if (root.TryGetProperty("changes", out var changes))
            {
                foreach (var change in changes.EnumerateArray())
                {
                    result.Changes.Add(new CellChange
                    {
                        Column = change.GetProperty("column").GetString()![0],
                        Row = change.GetProperty("row").GetInt32(),
                        Value = change.GetProperty("value").GetString() ?? string.Empty,
                        Formula = change.TryGetProperty("formula", out var f) && f.ValueKind != JsonValueKind.Null
                            ? f.GetString()
                            : null
                    });
                }
            }

            return result;
        }
        catch
        {
            return new SpreadsheetCommandResult
            {
                Success = false,
                Message = "応答の解析に失敗しました"
            };
        }
    }

    private DocumentCommandResult ParseDocumentResponse(string response)
    {
        try
        {
            var jsonText = ExtractJson(response);
            using var doc = JsonDocument.Parse(jsonText);
            var root = doc.RootElement;

            return new DocumentCommandResult
            {
                Success = root.GetProperty("success").GetBoolean(),
                Message = root.GetProperty("message").GetString() ?? string.Empty,
                NewText = root.TryGetProperty("newText", out var t) && t.ValueKind != JsonValueKind.Null
                    ? t.GetString()
                    : null
            };
        }
        catch
        {
            return new DocumentCommandResult
            {
                Success = false,
                Message = "応答の解析に失敗しました"
            };
        }
    }

    private MailCommandResult ParseMailResponse(string response)
    {
        try
        {
            var jsonText = ExtractJson(response);
            using var doc = JsonDocument.Parse(jsonText);
            var root = doc.RootElement;

            return new MailCommandResult
            {
                Success = root.GetProperty("success").GetBoolean(),
                Message = root.GetProperty("message").GetString() ?? string.Empty,
                Action = root.TryGetProperty("action", out var a) ? a.GetString() : null,
                To = root.TryGetProperty("to", out var t) ? t.GetString() : null,
                Subject = root.TryGetProperty("subject", out var s) ? s.GetString() : null,
                Body = root.TryGetProperty("body", out var b) ? b.GetString() : null
            };
        }
        catch
        {
            return new MailCommandResult
            {
                Success = false,
                Message = "応答の解析に失敗しました"
            };
        }
    }

    private string ExtractJson(string text)
    {
        // ```json ... ``` ブロックを探す
        var jsonStart = text.IndexOf("```json");
        if (jsonStart >= 0)
        {
            var start = text.IndexOf('\n', jsonStart) + 1;
            var end = text.IndexOf("```", start);
            if (end > start)
            {
                return text.Substring(start, end - start).Trim();
            }
        }

        // { で始まる部分を探す
        var braceStart = text.IndexOf('{');
        if (braceStart >= 0)
        {
            var braceEnd = text.LastIndexOf('}');
            if (braceEnd > braceStart)
            {
                return text.Substring(braceStart, braceEnd - braceStart + 1);
            }
        }

        return text;
    }

    public void Dispose()
    {
        if (!_disposed)
        {
            _httpClient.Dispose();
            _disposed = true;
        }
    }
}
