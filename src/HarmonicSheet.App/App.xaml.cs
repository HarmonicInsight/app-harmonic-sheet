using System.Windows;
using Microsoft.Extensions.DependencyInjection;
using HarmonicSheet.Services;
using HarmonicSheet.ViewModels;
using HarmonicSheet.Views;

namespace HarmonicSheet;

public partial class App : Application
{
    private ServiceProvider? _serviceProvider;

    public static IServiceProvider? Services { get; private set; }

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        // グローバル例外ハンドラー
        AppDomain.CurrentDomain.UnhandledException += (s, args) =>
        {
            var ex = args.ExceptionObject as Exception;
            MessageBox.Show(
                $"予期しないエラーが発生しました。\n\n{ex?.Message}\n\n{ex?.StackTrace}",
                "エラー",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        };

        try
        {
            // Syncfusionライセンスキーの登録（ある場合）
            // Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("YOUR_LICENSE_KEY");

            var services = new ServiceCollection();
            ConfigureServices(services);
            _serviceProvider = services.BuildServiceProvider();
            Services = _serviceProvider;

            // 初回起動時はチュートリアルを表示（コマンドライン引数で無効化可能）
            // チュートリアルを完全に無効化: --skip-tutorial
            var skipTutorial = e.Args.Contains("--skip-tutorial") || e.Args.Contains("--no-tutorial");

            if (!skipTutorial)
            {
                try
                {
                    var tutorialService = _serviceProvider.GetRequiredService<ITutorialService>();
                    if (!tutorialService.IsCompleted)
                    {
                        try
                        {
                            var tutorialWindow = new TutorialWindow(tutorialService);
                            tutorialWindow.ShowDialog();
                        }
                        catch (Exception tutEx)
                        {
                            // チュートリアルウィンドウ表示エラー
                            System.Diagnostics.Debug.WriteLine($"Tutorial window error: {tutEx.Message}");
                            MessageBox.Show(
                                $"チュートリアル表示でエラーが発生しましたが、アプリは起動します。\n\n{tutEx.Message}",
                                "チュートリアルエラー",
                                MessageBoxButton.OK,
                                MessageBoxImage.Warning);
                        }
                    }
                }
                catch (Exception tutServiceEx)
                {
                    // チュートリアルサービス取得エラー
                    System.Diagnostics.Debug.WriteLine($"Tutorial service error: {tutServiceEx.Message}");
                }
            }

            var mainWindow = _serviceProvider.GetRequiredService<MainWindow>();
            mainWindow.Show();
        }
        catch (Exception ex)
        {
            var errorMessage = $"アプリケーションの起動中にエラーが発生しました。\n\n";
            errorMessage += $"エラー: {ex.Message}\n\n";

            if (ex.InnerException != null)
            {
                errorMessage += $"詳細: {ex.InnerException.Message}\n\n";
            }

            errorMessage += $"スタックトレース:\n{ex.StackTrace}";

            MessageBox.Show(
                errorMessage,
                "起動エラー",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            Shutdown();
        }
    }

    private void ConfigureServices(IServiceCollection services)
    {
        // Core Services
        services.AddSingleton<ISpeechService, SpeechService>();
        services.AddSingleton<IClaudeService, ClaudeService>();
        services.AddSingleton<ISpreadsheetService, SpreadsheetService>();
        services.AddSingleton<IDocumentService, DocumentService>();
        services.AddSingleton<IMailService, MailService>();
        services.AddSingleton<IPrintService, PrintService>();

        // 拡張サービス
        services.AddSingleton<IContactService, ContactService>();
        services.AddSingleton<IAccessibilityService, AccessibilityService>();
        services.AddSingleton<ITutorialService, TutorialService>();

        // ViewModels
        services.AddTransient<MainViewModel>();
        services.AddTransient<SpreadsheetViewModel>();
        services.AddTransient<DocumentViewModel>();
        services.AddTransient<MailViewModel>();

        // Views
        services.AddTransient<MainWindow>();
    }

    protected override void OnExit(ExitEventArgs e)
    {
        // アクセシビリティ設定を保存
        if (_serviceProvider?.GetService<IAccessibilityService>() is AccessibilityService accessibility)
        {
            accessibility.SaveSettings();
        }

        _serviceProvider?.Dispose();
        base.OnExit(e);
    }
}
