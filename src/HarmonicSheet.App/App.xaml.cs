using System.Windows;
using Microsoft.Extensions.DependencyInjection;
using HarmonicSheet.Services;
using HarmonicSheet.ViewModels;

namespace HarmonicSheet;

public partial class App : Application
{
    private ServiceProvider? _serviceProvider;

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        var services = new ServiceCollection();
        ConfigureServices(services);
        _serviceProvider = services.BuildServiceProvider();

        var mainWindow = _serviceProvider.GetRequiredService<MainWindow>();
        mainWindow.Show();
    }

    private void ConfigureServices(IServiceCollection services)
    {
        // Services
        services.AddSingleton<ISpeechService, SpeechService>();
        services.AddSingleton<IClaudeService, ClaudeService>();
        services.AddSingleton<ISpreadsheetService, SpreadsheetService>();
        services.AddSingleton<IDocumentService, DocumentService>();
        services.AddSingleton<IMailService, MailService>();
        services.AddSingleton<IPrintService, PrintService>();

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
        _serviceProvider?.Dispose();
        base.OnExit(e);
    }
}
