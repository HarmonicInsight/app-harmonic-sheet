# HarmonicSheet

シニア向けのシンプルなオフィスアプリ - 文書・表計算・メールを1つに

## コンセプト

80代以上の高齢者でも迷わず使えるシンプルなオフィスアプリです。

### 特徴

- **大きな文字・ボタン** - 見やすく、押しやすい（サイズ調整可能）
- **3つの機能を1つに** - 文書、表、メールが1つのアプリで完結
- **音声入力対応** - 話すだけで文字入力
- **音声読み上げ** - 文書やメールを読み上げ
- **自然言語操作** - 「A2に1万円入れて」と言うだけで操作
- **印刷ワンクリック** - 簡単に印刷
- **連絡先帳** - よく使うメールアドレスを簡単管理
- **チュートリアル** - 初回起動時のガイド付き

## 機能

### 文書（ワード的機能）
- シンプルな文書作成
- Word形式(.docx)で保存・読み込み
- AIアシスタントによる文章作成支援
- 読み上げ機能

### 表（エクセル的機能）
- シンプルな表計算
- 自然言語での操作
  - 「A2に1万円入れて」
  - 「Aの3に5000円」
  - 「A1とA2とA3を足してA4に」
- Excel形式(.xlsx)で保存・読み込み

### メール
- シンプルなメール送受信
- 連絡先帳でよく使う人を登録
- メールの読み上げ機能
- AIによるメール文面作成支援

### アクセシビリティ
- **文字サイズ調整** - 70%〜150%で調整可能
- **読み上げ速度調整** - ゆっくり〜はやい
- **高コントラストモード** - より見やすい表示

### チュートリアル
- 初回起動時に使い方ガイドを表示
- 10ステップでアプリの使い方を説明
- いつでも設定画面から再表示可能

## 技術スタック

- .NET 8 / WPF
- Syncfusion WPF Controls
  - SfSpreadsheet（表計算）
  - SfRichTextBoxAdv（文書エディタ）
- MailKit（メール送受信）
- System.Speech（音声読み上げ）
- Claude API / Haiku（AIアシスタント）

## プロジェクト構造

```
app-harmonic-sheet/
├── HarmonicSheet.sln
├── README.md
├── insight-common/           # 共通ライブラリ
│   └── VoiceInputHelper.cs
└── src/HarmonicSheet.App/
    ├── App.xaml / .cs
    ├── MainWindow.xaml / .cs
    ├── SettingsWindow.xaml / .cs
    ├── Styles/
    │   └── SeniorFriendlyStyles.xaml
    ├── Views/
    │   ├── DocumentView.xaml / .cs
    │   ├── SpreadsheetView.xaml / .cs
    │   ├── MailView.xaml / .cs
    │   ├── ContactsWindow.xaml / .cs
    │   ├── ContactEditWindow.xaml / .cs
    │   └── TutorialWindow.xaml / .cs
    ├── ViewModels/
    │   ├── MainViewModel.cs
    │   ├── DocumentViewModel.cs
    │   ├── SpreadsheetViewModel.cs
    │   └── MailViewModel.cs
    ├── Services/
    │   ├── ClaudeService.cs
    │   ├── SpeechService.cs
    │   ├── ContactService.cs
    │   ├── AccessibilityService.cs
    │   ├── TutorialService.cs
    │   └── ...
    └── Models/
        ├── Contact.cs
        └── ...
```

## 必要要件

- Windows 10/11
- .NET 8 Runtime
- Syncfusion License (Community License available)

## セットアップ

```bash
# リポジトリをクローン
git clone https://github.com/HarmonicInsight/app-harmonic-sheet.git
cd app-harmonic-sheet

# 依存パッケージの復元
dotnet restore

# ビルド
dotnet build

# 実行
dotnet run --project src/HarmonicSheet.App
```

## 設定

### AIアシスタント
1. 設定画面を開く（右上の歯車アイコン）
2. Claude APIキーを入力
3. 保存

### メール
1. 設定画面を開く
2. メールアドレス、パスワード、サーバー情報を入力
3. 保存

### 文字サイズ
1. 設定画面を開く
2. 「文字の大きさ」スライダーで調整
3. プレビューで確認
4. 保存

## トラブルシューティング

### アプリが起動しない場合

**方法1: チュートリアルをスキップして起動**
```bash
dotnet run --project src/HarmonicSheet.App -- --skip-tutorial
```

**方法2: チュートリアル進捗ファイルを削除**
```powershell
# PowerShellで実行
Remove-Item "$env:APPDATA\HarmonicSheet\tutorial_progress.json" -ErrorAction SilentlyContinue
```

**方法3: ビルドキャッシュをクリア**
```bash
dotnet clean src/HarmonicSheet.App
dotnet build src/HarmonicSheet.App
dotnet run --project src/HarmonicSheet.App
```

### エラーメッセージの確認

起動時にエラーが発生する場合、詳細なエラーメッセージが表示されます。
エラーメッセージのスクリーンショットを撮って、[Issues](https://github.com/HarmonicInsight/app-harmonic-sheet/issues)に報告してください。

### よくある問題

**Q: Syncfusionのライセンス警告が表示される**
A: Syncfusionの[Community License](https://www.syncfusion.com/sales/communitylicense)を取得してください（無料）。
   ライセンスキーを取得したら、`App.xaml.cs`の22行目付近のコメントを外して設定してください。

**Q: 表計算で「表の初期化に失敗しました」と表示される**
A: Syncfusionパッケージが正しくインストールされているか確認してください:
```bash
dotnet restore
dotnet build
```

## ライセンス

MIT License

## 関連プロジェクト

- [InsightSheet](https://github.com/HarmonicInsight/app-Insight-excel) - コンサル・業務改善向けの高機能版
