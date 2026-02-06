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
        var helpMessage = @"📞 サポート窓口

■ 電話サポート
　📞 0120-XXX-XXX
　受付時間: 平日 9:00〜18:00

■ よくある質問
　・保存したファイルが見つからない
　　→「ホーム」タブから最近のファイルを確認

　・文字が小さい
　　→「設定」から文字サイズを変更できます

　・間違えて消してしまった
　　→「元に戻す」ボタンで元に戻せます

■ 遠隔サポート
　お電話いただければ、画面を見ながら
　サポートスタッフがご案内します

このメッセージをスクリーンショットして
ご家族に見せることもできます";

        var result = MessageBox.Show(
            helpMessage,
            "困ったときは - サポート窓口",
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
