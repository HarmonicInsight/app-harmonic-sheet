using System.Windows;
using System.Windows.Controls;
using HarmonicSheet.Models;
using HarmonicSheet.Services;

namespace HarmonicSheet.Views;

public partial class ContactsWindow : Window
{
    private readonly IContactService _contactService;
    private Contact? _selectedContact;

    /// <summary>選択された連絡先（ダイアログ結果）</summary>
    public Contact? SelectedContact => _selectedContact;

    public ContactsWindow(IContactService contactService)
    {
        InitializeComponent();
        _contactService = contactService;
        LoadContacts();
    }

    private void LoadContacts()
    {
        var contacts = _contactService.GetAllContacts();
        ContactsList.ItemsSource = contacts;
    }

    private void OnSearchTextChanged(object sender, TextChangedEventArgs e)
    {
        var query = SearchBox.Text;
        var contacts = string.IsNullOrWhiteSpace(query)
            ? _contactService.GetAllContacts()
            : _contactService.SearchContacts(query);

        ContactsList.ItemsSource = contacts;
    }

    private void OnGroupFilterChanged(object sender, SelectionChangedEventArgs e)
    {
        if (GroupFilter.SelectedItem is ComboBoxItem item)
        {
            var group = item.Content?.ToString();
            var contacts = group == "すべて"
                ? _contactService.GetAllContacts()
                : _contactService.GetContactsByGroup(group ?? "その他");

            ContactsList.ItemsSource = contacts;
        }
    }

    private void OnContactSelected(object sender, SelectionChangedEventArgs e)
    {
        _selectedContact = ContactsList.SelectedItem as Contact;
        var hasSelection = _selectedContact != null;

        EditButton.IsEnabled = hasSelection;
        DeleteButton.IsEnabled = hasSelection;
        SelectButton.IsEnabled = hasSelection;
    }

    private void OnAddContactClick(object sender, RoutedEventArgs e)
    {
        var dialog = new ContactEditWindow(null);
        dialog.Owner = this;

        if (dialog.ShowDialog() == true && dialog.Contact != null)
        {
            _contactService.AddContact(dialog.Contact);
            LoadContacts();
        }
    }

    private void OnEditContactClick(object sender, RoutedEventArgs e)
    {
        if (_selectedContact == null) return;

        var dialog = new ContactEditWindow(_selectedContact);
        dialog.Owner = this;

        if (dialog.ShowDialog() == true && dialog.Contact != null)
        {
            _contactService.UpdateContact(dialog.Contact);
            LoadContacts();
        }
    }

    private void OnDeleteContactClick(object sender, RoutedEventArgs e)
    {
        if (_selectedContact == null) return;

        var result = MessageBox.Show(
            $"「{_selectedContact.DisplayName}」を削除しますか？",
            "削除の確認",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question);

        if (result == MessageBoxResult.Yes)
        {
            _contactService.DeleteContact(_selectedContact.Id);
            _selectedContact = null;
            LoadContacts();
        }
    }

    private void OnSelectContactClick(object sender, RoutedEventArgs e)
    {
        if (_selectedContact != null)
        {
            _contactService.MarkAsUsed(_selectedContact.Id);
            DialogResult = true;
            Close();
        }
    }
}
