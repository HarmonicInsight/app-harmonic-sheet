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

        try
        {
            var services = new ServiceCollection();
            ConfigureServices(services);
            _serviceProvider = services.BuildServiceProvider();
            Services = _serviceProvider;

            // 初回起動時はチュートリアルを表示
            try
            {
                var tutorialService = _serviceProvider.GetRequiredService<ITutorialService>();
                if (!tutorialService.IsCompleted)
                {
                    var tutorialWindow = new TutorialWindow(tutorialService);
                    tutorialWindow.ShowDialog();
                }
            }
            catch
            {
                // チュートリアル表示エラーは無視して起動を続行
            }

            var mainWindow = _serviceProvider.GetRequiredService<MainWindow>();
            mainWindow.Show();
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"アプリケーションの起動中にエラーが発生しました。\n\n{ex.Message}",
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
