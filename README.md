# HarmonicSheet

シニア向けのシンプルなオフィスアプリ - 文書・表計算・メールを1つに

## コンセプト

80代以上の高齢者でも迷わず使えるシンプルなオフィスアプリです。

### 特徴

- **大きな文字・ボタン** - 見やすく、押しやすい
- **3つの機能を1つに** - 文書、表、メールが1つのアプリで完結
- **音声入力対応** - 話すだけで文字入力
- **音声読み上げ** - 文書やメールを読み上げ
- **自然言語操作** - 「A2に1万円入れて」と言うだけで操作
- **印刷ワンクリック** - 簡単に印刷

## 機能

### 文書（ワード的機能）
- シンプルな文書作成
- Word形式(.docx)で保存・読み込み
- AIアシスタントによる文章作成支援

### 表（エクセル的機能）
- シンプルな表計算
- 自然言語での操作
  - 「A2に1万円入れて」
  - 「A1とA2とA3を足してA4に」
- Excel形式(.xlsx)で保存・読み込み

### メール
- シンプルなメール送受信
- メールの読み上げ機能
- AIによるメール文面作成支援

## 技術スタック

- .NET 8 / WPF
- Syncfusion WPF Controls
  - SfSpreadsheet（表計算）
  - SfRichTextBoxAdv（文書エディタ）
- MailKit（メール送受信）
- System.Speech（音声読み上げ）
- Claude API / Haiku（AIアシスタント）

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

## ライセンス

MIT License

## 関連プロジェクト

- [InsightSheet](https://github.com/HarmonicInsight/app-Insight-excel) - コンサル・業務改善向けの高機能版
