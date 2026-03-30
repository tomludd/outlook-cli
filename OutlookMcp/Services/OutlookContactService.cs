using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace OutlookMcp.Services;

[SupportedOSPlatform("windows")]
public class OutlookContactService : IDisposable
{
    private const int OlFolderContacts = 10;
    private const int OlContactItem = 2;

    private dynamic? _outlookApp;

    private dynamic GetOutlookApp()
    {
        if (_outlookApp != null) return _outlookApp;

        var type = Type.GetTypeFromProgID("Outlook.Application")
            ?? throw new InvalidOperationException(
                "Microsoft Outlook is not installed or not registered on this system.");

        _outlookApp = Activator.CreateInstance(type)
            ?? throw new InvalidOperationException("Failed to create Outlook.Application instance.");

        return _outlookApp;
    }

    private dynamic GetNamespace() => GetOutlookApp().GetNamespace("MAPI");

    private dynamic GetStoreFolder(string? account, int folderType)
    {
        var ns = GetNamespace();

        if (string.IsNullOrEmpty(account))
            return ns.GetDefaultFolder(folderType);

        var stores = ns.Stores;
        for (int i = 1; i <= stores.Count; i++)
        {
            var store = stores.Item(i);
            if (string.Equals((string)store.DisplayName, account, StringComparison.OrdinalIgnoreCase))
                return store.GetDefaultFolder(folderType);
        }

        throw new InvalidOperationException($"Account not found: {account}. Use list_accounts to see available accounts.");
    }

    public List<Dictionary<string, object?>> ListContacts(int count, string? account = null)
    {
        var ns = GetNamespace();
        var folder = GetStoreFolder(account, OlFolderContacts);
        var items = folder.Items;
        items.Sort("[LastName]");

        var contacts = new List<Dictionary<string, object?>>();
        int limit = Math.Min(count, items.Count);
        for (int i = 1; i <= limit; i++)
        {
            try
            {
                var item = items.Item(i);
                if ((int)item.Class == 40) // olContact
                    contacts.Add(ContactToDict(item));
            }
            catch
            {
                // Skip non-contact items (distribution lists, etc.)
            }
        }
        return contacts;
    }

    public List<Dictionary<string, object?>> SearchContacts(string query, int maxResults, string? account = null)
    {
        var ns = GetNamespace();
        var folder = GetStoreFolder(account, OlFolderContacts);

        var filter = $"@SQL=(\"urn:schemas:contacts:cn\" LIKE '%{EscapeDasl(query)}%' " +
                     $"OR \"urn:schemas:contacts:email1\" LIKE '%{EscapeDasl(query)}%' " +
                     $"OR \"urn:schemas:contacts:o\" LIKE '%{EscapeDasl(query)}%')";

        var items = folder.Items.Restrict(filter);

        var contacts = new List<Dictionary<string, object?>>();
        int limit = Math.Min(maxResults, items.Count);
        for (int i = 1; i <= limit; i++)
        {
            try
            {
                contacts.Add(ContactToDict(items.Item(i)));
            }
            catch
            {
                // Skip non-contact items
            }
        }
        return contacts;
    }

    public Dictionary<string, object?> GetContact(string entryId)
    {
        var ns = GetNamespace();
        dynamic item;
        try
        {
            item = ns.GetItemFromID(entryId);
        }
        catch
        {
            throw new InvalidOperationException($"Contact not found with ID: {entryId}");
        }
        return ContactToDict(item);
    }

    public string CreateContact(string? firstName, string? lastName, string? email,
        string? phone, string? mobilePhone, string? company, string? jobTitle,
        string? businessAddress, string? notes, string? account = null)
    {
        var folder = GetStoreFolder(account, OlFolderContacts);
        var contact = folder.Items.Add(OlContactItem);

        if (!string.IsNullOrEmpty(firstName)) contact.FirstName = firstName;
        if (!string.IsNullOrEmpty(lastName)) contact.LastName = lastName;
        if (!string.IsNullOrEmpty(email)) contact.Email1Address = email;
        if (!string.IsNullOrEmpty(phone)) contact.BusinessTelephoneNumber = phone;
        if (!string.IsNullOrEmpty(mobilePhone)) contact.MobileTelephoneNumber = mobilePhone;
        if (!string.IsNullOrEmpty(company)) contact.CompanyName = company;
        if (!string.IsNullOrEmpty(jobTitle)) contact.JobTitle = jobTitle;
        if (!string.IsNullOrEmpty(businessAddress)) contact.BusinessAddress = businessAddress;
        if (!string.IsNullOrEmpty(notes)) contact.Body = notes;

        contact.Save();

        string entryId = contact.EntryID;
        Marshal.ReleaseComObject(contact);
        return entryId;
    }

    public bool UpdateContact(string entryId, string? firstName, string? lastName,
        string? email, string? phone, string? mobilePhone, string? company,
        string? jobTitle, string? businessAddress, string? notes)
    {
        var ns = GetNamespace();
        dynamic contact;
        try
        {
            contact = ns.GetItemFromID(entryId);
        }
        catch
        {
            throw new InvalidOperationException($"Contact not found with ID: {entryId}");
        }

        if (!string.IsNullOrEmpty(firstName)) contact.FirstName = firstName;
        if (!string.IsNullOrEmpty(lastName)) contact.LastName = lastName;
        if (!string.IsNullOrEmpty(email)) contact.Email1Address = email;
        if (!string.IsNullOrEmpty(phone)) contact.BusinessTelephoneNumber = phone;
        if (!string.IsNullOrEmpty(mobilePhone)) contact.MobileTelephoneNumber = mobilePhone;
        if (!string.IsNullOrEmpty(company)) contact.CompanyName = company;
        if (!string.IsNullOrEmpty(jobTitle)) contact.JobTitle = jobTitle;
        if (!string.IsNullOrEmpty(businessAddress)) contact.BusinessAddress = businessAddress;
        if (!string.IsNullOrEmpty(notes)) contact.Body = notes;

        contact.Save();
        Marshal.ReleaseComObject(contact);
        return true;
    }

    public bool DeleteContact(string entryId)
    {
        var ns = GetNamespace();
        dynamic contact;
        try
        {
            contact = ns.GetItemFromID(entryId);
        }
        catch
        {
            throw new InvalidOperationException($"Contact not found with ID: {entryId}");
        }

        contact.Delete();
        Marshal.ReleaseComObject(contact);
        return true;
    }

    private static Dictionary<string, object?> ContactToDict(dynamic contact)
    {
        return new Dictionary<string, object?>
        {
            ["id"] = (string)contact.EntryID,
            ["fullName"] = SafeGet(() => (string)contact.FullName),
            ["firstName"] = SafeGet(() => (string)contact.FirstName),
            ["lastName"] = SafeGet(() => (string)contact.LastName),
            ["email"] = SafeGet(() => (string)contact.Email1Address),
            ["phone"] = SafeGet(() => (string)contact.BusinessTelephoneNumber),
            ["mobilePhone"] = SafeGet(() => (string)contact.MobileTelephoneNumber),
            ["company"] = SafeGet(() => (string)contact.CompanyName),
            ["jobTitle"] = SafeGet(() => (string)contact.JobTitle),
            ["businessAddress"] = SafeGet(() => (string)contact.BusinessAddress),
        };
    }

    private static string? SafeGet(Func<string> getter)
    {
        try
        {
            var val = getter();
            return string.IsNullOrEmpty(val) ? null : val;
        }
        catch
        {
            return null;
        }
    }

    private static string EscapeDasl(string value) =>
        value.Replace("'", "''").Replace("\"", "\"\"");

    public void Dispose()
    {
        if (_outlookApp != null)
        {
            Marshal.ReleaseComObject(_outlookApp);
            _outlookApp = null;
        }
    }
}
