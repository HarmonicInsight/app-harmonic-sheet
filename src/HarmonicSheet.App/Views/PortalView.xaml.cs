using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace HarmonicSheet.Views;

public partial class PortalView : UserControl
{
    private static readonly string RecentSpreadsheetPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "HarmonicSheet", "recent_spreadsheets.txt");

    private static readonly string RecentDocumentPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "HarmonicSheet", "recent_documents.txt");

    public event EventHandler<string>? SpreadsheetSelected;
    public event EventHandler<string>? DocumentSelected;
    public event EventHandler<string>? MailSelected;

    public PortalView()
    {
        InitializeComponent();
        Loaded += OnLoaded;
    }

    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        LoadRecentFiles();
    }

    public void LoadRecentFiles()
    {
        // Load recent spreadsheets
        var spreadsheets = new ObservableCollection<RecentFileItem>();
        try
        {
            if (File.Exists(RecentSpreadsheetPath))
            {
                var files = File.ReadAllLines(RecentSpreadsheetPath)
                    .Where(f => !string.IsNullOrWhiteSpace(f) && File.Exists(f))
                    .Take(5);

                foreach (var file in files)
                {
                    var fileInfo = new FileInfo(file);
                    spreadsheets.Add(new RecentFileItem
                    {
                        Path = file,
                        FileName = Path.GetFileName(file),
                        LastAccessed = fileInfo.LastAccessTime
                    });
                }
            }
        }
        catch { }

        if (spreadsheets.Count == 0)
        {
            spreadsheets.Add(new RecentFileItem
            {
                Path = "",
                FileName = "まだファイルがありません",
                LastAccessed = DateTime.Now
            });
        }
        RecentSpreadsheets.ItemsSource = spreadsheets;

        // Load recent documents
        var documents = new ObservableCollection<RecentFileItem>();
        try
        {
            if (File.Exists(RecentDocumentPath))
            {
                var files = File.ReadAllLines(RecentDocumentPath)
                    .Where(f => !string.IsNullOrWhiteSpace(f) && File.Exists(f))
                    .Take(5);

                foreach (var file in files)
                {
                    var fileInfo = new FileInfo(file);
                    documents.Add(new RecentFileItem
                    {
                        Path = file,
                        FileName = Path.GetFileName(file),
                        LastAccessed = fileInfo.LastAccessTime
                    });
                }
            }
        }
        catch { }

        if (documents.Count == 0)
        {
            documents.Add(new RecentFileItem
            {
                Path = "",
                FileName = "まだファイルがありません",
                LastAccessed = DateTime.Now
            });
        }
        RecentDocuments.ItemsSource = documents;

        // Load recent mails (mock data for now)
        var mails = new ObservableCollection<RecentMailItem>
        {
            new RecentMailItem
            {
                Id = "1",
                Subject = "メール機能は準備中です",
                From = "システム",
                ReceivedDate = DateTime.Now
            }
        };
        RecentMails.ItemsSource = mails;
    }

    private void OnOpenSpreadsheet(object sender, RoutedEventArgs e)
    {
        if (sender is Button button && button.Tag is string path && !string.IsNullOrEmpty(path))
        {
            SpreadsheetSelected?.Invoke(this, path);
        }
    }

    private void OnOpenDocument(object sender, RoutedEventArgs e)
    {
        if (sender is Button button && button.Tag is string path && !string.IsNullOrEmpty(path))
        {
            DocumentSelected?.Invoke(this, path);
        }
    }

    private void OnOpenMail(object sender, RoutedEventArgs e)
    {
        if (sender is Button button && button.Tag is string id)
        {
            MailSelected?.Invoke(this, id);
        }
    }

    public class RecentFileItem
    {
        public string Path { get; set; } = "";
        public string FileName { get; set; } = "";
        public DateTime LastAccessed { get; set; }
    }

    public class RecentMailItem
    {
        public string Id { get; set; } = "";
        public string Subject { get; set; } = "";
        public string From { get; set; } = "";
        public DateTime ReceivedDate { get; set; }
    }
}
