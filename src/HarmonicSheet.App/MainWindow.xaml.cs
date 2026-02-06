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
        if (DocumentView == null || SpreadsheetView == null || MailView == null || StatusText == null)
            return;

        // 全てのビューを非表示に
        DocumentView.Visibility = Visibility.Collapsed;
        SpreadsheetView.Visibility = Visibility.Collapsed;
        MailView.Visibility = Visibility.Collapsed;

        // 選択されたタブのビューを表示
        if (TabDocument.IsChecked == true)
        {
            DocumentView.Visibility = Visibility.Visible;
            StatusText.Text = "文書モード - 文章を書きましょう";
        }
        else if (TabSpreadsheet.IsChecked == true)
        {
            SpreadsheetView.Visibility = Visibility.Visible;
            StatusText.Text = "表モード - 数字を入れたり計算できます";
        }
        else if (TabMail.IsChecked == true)
        {
            MailView.Visibility = Visibility.Visible;
            StatusText.Text = "メールモード - メールを送ったり読んだりできます";
        }
    }

    private void OnVoiceInputClick(object sender, RoutedEventArgs e)
    {
        // Windows音声入力を起動（Win+H）
        _speechService.ActivateWindowsVoiceTyping();
        StatusText.Text = "音声入力中... 話し終わったらもう一度クリック";
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
