using System.IO;
using System.Text.Json;

namespace HarmonicSheet.Services;

/// <summary>
/// チュートリアルサービス
/// 初回起動時や操作のガイドを提供
/// </summary>
public interface ITutorialService
{
    /// <summary>チュートリアルが完了済みか</summary>
    bool IsCompleted { get; }

    /// <summary>現在のステップ</summary>
    int CurrentStep { get; }

    /// <summary>総ステップ数</summary>
    int TotalSteps { get; }

    /// <summary>現在のステップの情報を取得</summary>
    TutorialStep? GetCurrentStep();

    /// <summary>次のステップへ</summary>
    void NextStep();

    /// <summary>前のステップへ</summary>
    void PreviousStep();

    /// <summary>チュートリアルをスキップ</summary>
    void Skip();

    /// <summary>チュートリアルをリセット</summary>
    void Reset();

    /// <summary>特定のトピックのヒントを取得</summary>
    string GetHint(string topic);
}

/// <summary>
/// チュートリアルの1ステップ
/// </summary>
public class TutorialStep
{
    public int StepNumber { get; set; }
    public string Title { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public string? TargetElement { get; set; }
    public string? ImagePath { get; set; }
}

public class TutorialService : ITutorialService
{
    private static readonly string ProgressFilePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "HarmonicSheet",
        "tutorial_progress.json");

    private int _currentStep = 0;
    private bool _isCompleted = false;

    private readonly List<TutorialStep> _steps = new()
    {
        new TutorialStep
        {
            StepNumber = 1,
            Title = "HarmonicSheetへようこそ！",
            Description = "このアプリは「文書」「表」「メール」の\n3つの機能が1つになっています。\n\n画面上の大きなボタンで切り替えられます。",
            TargetElement = "TabButtons"
        },
        new TutorialStep
        {
            StepNumber = 2,
            Title = "文書を作る",
            Description = "「文書」タブでは手紙や報告書を作れます。\n\n文字を大きく表示して、\n読みやすくしています。",
            TargetElement = "DocumentTab"
        },
        new TutorialStep
        {
            StepNumber = 3,
            Title = "表を作る",
            Description = "「表」タブでは数字を入れて計算できます。\n\n例えば家計簿や名簿を作るときに使います。",
            TargetElement = "SpreadsheetTab"
        },
        new TutorialStep
        {
            StepNumber = 4,
            Title = "話しかけて操作する",
            Description = "表は話しかけるだけで操作できます。\n\n例：\n「A2に1万円入れて」\n「A1とA2を足してA3に」\n\n下の入力欄に打ち込んでください。",
            TargetElement = "CommandInput"
        },
        new TutorialStep
        {
            StepNumber = 5,
            Title = "声で文字を入力する",
            Description = "画面下の赤い丸ボタン（マイク）を押すと\n声で文字を入力できます。\n\n話し終わったら自動で文字になります。",
            TargetElement = "VoiceButton"
        },
        new TutorialStep
        {
            StepNumber = 6,
            Title = "メールを送る・読む",
            Description = "「メール」タブでメールを送ったり\n届いたメールを読んだりできます。\n\n連絡先帳によく使う人を登録しておくと\n宛先を選ぶだけで送れます。",
            TargetElement = "MailTab"
        },
        new TutorialStep
        {
            StepNumber = 7,
            Title = "読み上げ機能",
            Description = "メールや文書を声で読み上げることができます。\n\n「読み上げ」ボタンを押すと\nゆっくり読んでくれます。",
            TargetElement = "ReadAloudButton"
        },
        new TutorialStep
        {
            StepNumber = 8,
            Title = "印刷する",
            Description = "各画面の「印刷」ボタンを押すと\n紙に印刷できます。\n\n大きな文字で印刷されるので\n読みやすいです。",
            TargetElement = "PrintButton"
        },
        new TutorialStep
        {
            StepNumber = 9,
            Title = "設定を変える",
            Description = "右上の歯車ボタンで設定画面を開けます。\n\n・文字の大きさを変える\n・読み上げの速さを変える\n・メールの設定をする\n\nなどができます。",
            TargetElement = "SettingsButton"
        },
        new TutorialStep
        {
            StepNumber = 10,
            Title = "準備完了！",
            Description = "これで基本的な使い方は終わりです。\n\n困ったときは右上の「？」ボタンを\n押してください。\n\nそれでは、お楽しみください！",
            TargetElement = "HelpButton"
        }
    };

    private readonly Dictionary<string, string> _hints = new()
    {
        ["spreadsheet_input"] = "セルに値を入れるには:\n「A1に100を入れて」\n「Bの3に5000円」\nのように話しかけてください。",
        ["spreadsheet_formula"] = "計算するには:\n「A1とA2を足してA3に」\n「A列の合計をA10に」\nのように話しかけてください。",
        ["mail_compose"] = "メールを送るには:\n1. 「新しいメール」を押す\n2. 宛先を選ぶか入力\n3. 内容を書く\n4. 「送信」を押す",
        ["mail_reply"] = "返信するには:\nメールを選んで「返信」を押してください。\n宛先と件名は自動で入ります。",
        ["voice_input"] = "声で入力するには:\n1. 赤いマイクボタンを押す\n2. ゆっくりはっきり話す\n3. 話し終わると自動で文字になる",
        ["print"] = "印刷するには:\n「印刷」ボタンを押すだけです。\nプリンターの電源が入っているか\n確認してください。",
        ["save"] = "保存するには:\n「保存」ボタンを押して\nファイル名を入れてください。\nデスクトップに保存すると見つけやすいです。"
    };

    public TutorialService()
    {
        LoadProgress();
    }

    public bool IsCompleted => _isCompleted;
    public int CurrentStep => _currentStep;
    public int TotalSteps => _steps.Count;

    public TutorialStep? GetCurrentStep()
    {
        if (_isCompleted || _currentStep < 0 || _currentStep >= _steps.Count)
            return null;

        return _steps[_currentStep];
    }

    public void NextStep()
    {
        if (_currentStep < _steps.Count - 1)
        {
            _currentStep++;
            SaveProgress();
        }
        else
        {
            _isCompleted = true;
            SaveProgress();
        }
    }

    public void PreviousStep()
    {
        if (_currentStep > 0)
        {
            _currentStep--;
            SaveProgress();
        }
    }

    public void Skip()
    {
        _isCompleted = true;
        SaveProgress();
    }

    public void Reset()
    {
        _currentStep = 0;
        _isCompleted = false;
        SaveProgress();
    }

    public string GetHint(string topic)
    {
        return _hints.TryGetValue(topic, out var hint)
            ? hint
            : "ヘルプが見つかりませんでした。\n設定画面の「ヘルプ」を確認してください。";
    }

    private void LoadProgress()
    {
        try
        {
            if (File.Exists(ProgressFilePath))
            {
                var json = File.ReadAllText(ProgressFilePath);
                var progress = JsonSerializer.Deserialize<TutorialProgress>(json);
                if (progress != null)
                {
                    _currentStep = progress.CurrentStep;
                    _isCompleted = progress.IsCompleted;
                }
            }
        }
        catch
        {
            // 読み込みエラーは無視（初回起動として扱う）
        }
    }

    private void SaveProgress()
    {
        try
        {
            var directory = Path.GetDirectoryName(ProgressFilePath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            var progress = new TutorialProgress
            {
                CurrentStep = _currentStep,
                IsCompleted = _isCompleted
            };

            var json = JsonSerializer.Serialize(progress);
            File.WriteAllText(ProgressFilePath, json);
        }
        catch
        {
            // 保存エラーは無視
        }
    }

    private class TutorialProgress
    {
        public int CurrentStep { get; set; }
        public bool IsCompleted { get; set; }
    }
}
