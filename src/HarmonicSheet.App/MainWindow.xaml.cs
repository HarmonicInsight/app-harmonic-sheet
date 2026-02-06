using System.Windows;
using System.Windows.Controls;
using HarmonicSheet.ViewModels;
using HarmonicSheet.Services;

namespace HarmonicSheet;

public partial class MainWindow : Window
{
    private readonly MainViewModel _viewModel;
    private readonly ISpeechService _speechService;

    public MainWindow(MainViewModel viewModel, ISpeechService speechService)
    {
        InitializeComponent();
        _viewModel = viewModel;
        _speechService = speechService;
        DataContext = _viewModel;
    }

    private void OnTabChanged(object sender, RoutedEventArgs e)
    {
        // 初期化が完了していない場合は何もしない
        if (PortalView == null || DocumentView == null || SpreadsheetView == null || MailView == null)
            return;

        // 全てのビューを非表示に
        PortalView.Visibility = Visibility.Collapsed;
        DocumentView.Visibility = Visibility.Collapsed;
        SpreadsheetView.Visibility = Visibility.Collapsed;
        MailView.Visibility = Visibility.Collapsed;

        // 選択されたタブのビューを表示
        if (TabPortal.IsChecked == true)
        {
            PortalView.Visibility = Visibility.Visible;
            PortalView.LoadRecentFiles(); // ポータル表示時にファイルリストを更新
        }
        else if (TabSpreadsheet.IsChecked == true)
        {
            SpreadsheetView.Visibility = Visibility.Visible;
        }
        else if (TabDocument.IsChecked == true)
        {
            DocumentView.Visibility = Visibility.Visible;
        }
        else if (TabMail.IsChecked == true)
        {
            MailView.Visibility = Visibility.Visible;
        }
    }

    private void OnHelpLineClick(object sender, RoutedEventArgs e)
    {
        var helpMessage = @"🆘 困ったとき - AIアシスタント

━━━━━━━━━━━━━━━━━━━━━━

🤖 何かお困りですか？

💬 よくある質問（すぐに解決）

Q1: ファイルが見つからない
→ 「ホーム」ボタンを押して、
   最近使ったファイルを確認してください

Q2: 文字が小さくて見えない
→ 「⚙ 設定」ボタンから
   文字の大きさを変更できます

Q3: 間違えて消してしまった
→ 「元に戻す」ボタン（↩）で
   元に戻せます

Q4: 使い方がわからない
→ 各機能のボタンにマウスを合わせると
   説明が表示されます

━━━━━━━━━━━━━━━━━━━━━━

📞 電話で直接サポート
　0120-XXX-XXX
　受付: 平日 9:00〜18:00

　お電話いただければ、画面を見ながら
　スタッフが丁寧にご案内します

━━━━━━━━━━━━━━━━━━━━━━

💡 このメッセージをスクリーンショット
   して、ご家族に見せることもできます";

        var result = MessageBox.Show(
            helpMessage,
            "🆘 困ったとき - AIアシスタント",
            MessageBoxButton.OK,
            MessageBoxImage.Information);
    }

    private void OnSettingsClick(object sender, RoutedEventArgs e)
    {
        var settingsWindow = new SettingsWindow();
        settingsWindow.Owner = this;
        settingsWindow.ShowDialog();
    }

    private void OnHelpClick(object sender, RoutedEventArgs e)
    {
        var helpText = @"【HarmonicOffice の使い方】

■ 文書タブ
　文章を書くことができます。
　手紙や報告書を作るときに使います。

■ 表タブ
　数字を入れて計算ができます。
　家計簿や名簿を作るときに使います。

　「A1に1000円入れて」のように
　話しかけるだけで操作できます。

■ メールタブ
　メールを送ったり読んだりできます。

■ 音声入力
　画面下の赤い丸ボタンを押すと
　声で文字を入力できます。

■ 印刷
　各画面の「印刷」ボタンで印刷できます。";

        MessageBox.Show(helpText, "ヘルプ", MessageBoxButton.OK, MessageBoxImage.Information);
    }
}
