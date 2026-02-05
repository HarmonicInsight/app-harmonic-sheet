using System.IO;
using System.Text.Json;
using HarmonicSheet.Models;

namespace HarmonicSheet.Services;

/// <summary>
/// 連絡先管理サービス
/// </summary>
public interface IContactService
{
    List<Contact> GetAllContacts();
    List<Contact> GetContactsByGroup(string group);
    List<Contact> GetFavoriteContacts();
    List<Contact> SearchContacts(string query);
    Contact? GetContactById(string id);
    Contact? GetContactByEmail(string email);
    void AddContact(Contact contact);
    void UpdateContact(Contact contact);
    void DeleteContact(string id);
    void MarkAsUsed(string id);
    void SaveContacts();
}

public class ContactService : IContactService
{
    private readonly List<Contact> _contacts = new();
    private static readonly string ContactsFilePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "HarmonicSheet",
        "contacts.json");

    public ContactService()
    {
        LoadContacts();
    }

    private void LoadContacts()
    {
        try
        {
            if (File.Exists(ContactsFilePath))
            {
                var json = File.ReadAllText(ContactsFilePath);
                var contacts = JsonSerializer.Deserialize<List<Contact>>(json);
                if (contacts != null)
                {
                    _contacts.Clear();
                    _contacts.AddRange(contacts);
                }
            }
        }
        catch
        {
            // 読み込みエラーは無視
        }
    }

    public void SaveContacts()
    {
        try
        {
            var directory = Path.GetDirectoryName(ContactsFilePath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            var json = JsonSerializer.Serialize(_contacts, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(ContactsFilePath, json);
        }
        catch
        {
            // 保存エラーは無視
        }
    }

    public List<Contact> GetAllContacts()
    {
        return _contacts
            .OrderByDescending(c => c.IsFavorite)
            .ThenByDescending(c => c.LastUsed)
            .ThenBy(c => c.DisplayName)
            .ToList();
    }

    public List<Contact> GetContactsByGroup(string group)
    {
        return _contacts
            .Where(c => c.Group == group)
            .OrderByDescending(c => c.IsFavorite)
            .ThenByDescending(c => c.LastUsed)
            .ToList();
    }

    public List<Contact> GetFavoriteContacts()
    {
        return _contacts
            .Where(c => c.IsFavorite)
            .OrderByDescending(c => c.LastUsed)
            .ToList();
    }

    public List<Contact> SearchContacts(string query)
    {
        if (string.IsNullOrWhiteSpace(query))
            return GetAllContacts();

        var lowerQuery = query.ToLowerInvariant();
        return _contacts
            .Where(c =>
                c.DisplayName.ToLowerInvariant().Contains(lowerQuery) ||
                c.Email.ToLowerInvariant().Contains(lowerQuery) ||
                (c.Notes?.ToLowerInvariant().Contains(lowerQuery) ?? false))
            .OrderByDescending(c => c.IsFavorite)
            .ThenByDescending(c => c.LastUsed)
            .ToList();
    }

    public Contact? GetContactById(string id)
    {
        return _contacts.FirstOrDefault(c => c.Id == id);
    }

    public Contact? GetContactByEmail(string email)
    {
        return _contacts.FirstOrDefault(c =>
            c.Email.Equals(email, StringComparison.OrdinalIgnoreCase));
    }

    public void AddContact(Contact contact)
    {
        _contacts.Add(contact);
        SaveContacts();
    }

    public void UpdateContact(Contact contact)
    {
        var existing = _contacts.FirstOrDefault(c => c.Id == contact.Id);
        if (existing != null)
        {
            existing.DisplayName = contact.DisplayName;
            existing.Email = contact.Email;
            existing.PhoneNumber = contact.PhoneNumber;
            existing.Notes = contact.Notes;
            existing.Group = contact.Group;
            existing.IsFavorite = contact.IsFavorite;
            SaveContacts();
        }
    }

    public void DeleteContact(string id)
    {
        var contact = _contacts.FirstOrDefault(c => c.Id == id);
        if (contact != null)
        {
            _contacts.Remove(contact);
            SaveContacts();
        }
    }

    public void MarkAsUsed(string id)
    {
        var contact = _contacts.FirstOrDefault(c => c.Id == id);
        if (contact != null)
        {
            contact.LastUsed = DateTime.Now;
            SaveContacts();
        }
    }
}
