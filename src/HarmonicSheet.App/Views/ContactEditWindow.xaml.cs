using System.Windows;
using System.Windows.Controls;
using HarmonicSheet.Models;

namespace HarmonicSheet.Views;

public partial class ContactEditWindow : Window
{
    private readonly Contact? _existingContact;

    /// <summary>編集後の連絡先</summary>
    public Contact? Contact { get; private set; }

    public ContactEditWindow(Contact? existingContact)
    {
        InitializeComponent();
        _existingContact = existingContact;

        if (existingContact != null)
        {
            Title = "連絡先の編集";
            LoadContact(existingContact);
        }
        else
        {
            Title = "新しい連絡先";
        }
    }

    private void LoadContact(Contact contact)
    {
        NameInput.Text = contact.DisplayName;
        EmailInput.Text = contact.Email;
        PhoneInput.Text = contact.PhoneNumber ?? string.Empty;
        NotesInput.Text = contact.Notes ?? string.Empty;
        FavoriteCheck.IsChecked = contact.IsFavorite;

        // グループを選択
        foreach (ComboBoxItem item in GroupInput.Items)
        {
            if (item.Content?.ToString() == contact.Group)
            {
                GroupInput.SelectedItem = item;
                break;
            }
        }
    }

    private void OnSaveClick(object sender, RoutedEventArgs e)
    {
        // バリデーション
        if (string.IsNullOrWhiteSpace(NameInput.Text))
        {
            MessageBox.Show("名前を入力してください。", "入力エラー",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            NameInput.Focus();
            return;
        }

        if (string.IsNullOrWhiteSpace(EmailInput.Text))
        {
            MessageBox.Show("メールアドレスを入力してください。", "入力エラー",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            EmailInput.Focus();
            return;
        }

        // メールアドレスの簡易チェック
        if (!EmailInput.Text.Contains("@"))
        {
            MessageBox.Show("正しいメールアドレスを入力してください。", "入力エラー",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            EmailInput.Focus();
            return;
        }

        // 連絡先を作成/更新
        Contact = _existingContact ?? new Contact();
        Contact.DisplayName = NameInput.Text.Trim();
        Contact.Email = EmailInput.Text.Trim();
        Contact.PhoneNumber = string.IsNullOrWhiteSpace(PhoneInput.Text) ? null : PhoneInput.Text.Trim();
        Contact.Notes = string.IsNullOrWhiteSpace(NotesInput.Text) ? null : NotesInput.Text.Trim();
        Contact.IsFavorite = FavoriteCheck.IsChecked ?? false;
        Contact.Group = (GroupInput.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "その他";

        DialogResult = true;
        Close();
    }

    private void OnCancelClick(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }
}
