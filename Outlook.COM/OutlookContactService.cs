using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace Outlook.COM;

[SupportedOSPlatform("windows")]
public class OutlookContactService
{
    private const int OlFolderContacts = 10;
    private const int OlContactItem = 2;

    private dynamic GetOutlookApp() => OutlookComHost.GetApp();

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
        => OutlookComInvoker.Run(() => ListContactsCore(count, account));

    private List<Dictionary<string, object?>> ListContactsCore(int count, string? account)
    {
        if (!string.IsNullOrEmpty(account))
            return CollectContacts(GetStoreFolder(account, OlFolderContacts), count, account);

        // Aggregate across all accounts
        var ns = GetNamespace();
        var all = new List<Dictionary<string, object?>>();
        var stores = ns.Stores;
        try
        {
            for (int i = 1; i <= stores.Count; i++)
            {
                dynamic? store = null;
                try { store = stores.Item(i); all.AddRange(CollectContacts(store.GetDefaultFolder(OlFolderContacts), count, (string)store.DisplayName)); }
                catch { /* Store may not have a contacts folder */ }
                finally { OutlookComHost.Release(store); }
            }
        }
        finally
        {
            OutlookComHost.Release(stores);
            OutlookComHost.Release(ns);
        }
        all.Sort((a, b) => string.Compare(a["fullName"]?.ToString(), b["fullName"]?.ToString(), StringComparison.OrdinalIgnoreCase));
        return all.Take(count).ToList();
    }

    private static List<Dictionary<string, object?>> CollectContacts(dynamic folder, int count, string? accountName)
    {
        var items = folder.Items;
        items.Sort("[LastName]");
        var contacts = new List<Dictionary<string, object?>>();
        int limit = Math.Min(count, items.Count);
        try
        {
            for (int i = 1; i <= limit; i++)
            {
                dynamic? item = null;
                try
                {
                    item = items.Item(i);
                    if ((int)item.Class == 40) // olContact
                    {
                        var contact = ContactToDict(item);
                        if (accountName != null) contact["account"] = accountName;
                        contacts.Add(contact);
                    }
                }
                catch { /* Skip non-contact items (distribution lists, etc.) */ }
                finally { OutlookComHost.Release(item); }
            }
        }
        finally { OutlookComHost.Release(items); }
        return contacts;
    }

    public List<Dictionary<string, object?>> SearchContacts(string query, int maxResults, string? account = null)
        => OutlookComInvoker.Run(() => SearchContactsCore(query, maxResults, account));

    private List<Dictionary<string, object?>> SearchContactsCore(string query, int maxResults, string? account)
    {
        var filter = $"@SQL=(\"urn:schemas:contacts:cn\" LIKE '%{EscapeDasl(query)}%' " +
                     $"OR \"urn:schemas:contacts:email1\" LIKE '%{EscapeDasl(query)}%' " +
                     $"OR \"urn:schemas:contacts:o\" LIKE '%{EscapeDasl(query)}%')";

        List<Dictionary<string, object?>> SearchFolder(dynamic folder, string? accountName)
        {
            var items = folder.Items.Restrict(filter);
            var results = new List<Dictionary<string, object?>>();
            int limit = Math.Min(maxResults, items.Count);
            try
            {
                for (int i = 1; i <= limit; i++)
                {
                    dynamic? item = null;
                    try
                    {
                        item = items.Item(i);
                        var contact = ContactToDict(item);
                        if (accountName != null) contact["account"] = accountName;
                        results.Add(contact);
                    }
                    catch { /* Skip non-contact items */ }
                    finally { OutlookComHost.Release(item); }
                }
            }
            finally { OutlookComHost.Release(items); }
            return results;
        }

        if (!string.IsNullOrEmpty(account))
            return SearchFolder(GetStoreFolder(account, OlFolderContacts), account);

        // Search across all accounts
        var ns = GetNamespace();
        var all = new List<Dictionary<string, object?>>();
        var stores = ns.Stores;
        try
        {
            for (int i = 1; i <= stores.Count; i++)
            {
                dynamic? store = null;
                try { store = stores.Item(i); all.AddRange(SearchFolder(store.GetDefaultFolder(OlFolderContacts), (string)store.DisplayName)); }
                catch { /* Store may not have a contacts folder */ }
                finally { OutlookComHost.Release(store); }
            }
        }
        finally
        {
            OutlookComHost.Release(stores);
            OutlookComHost.Release(ns);
        }
        return all.Take(maxResults).ToList();
    }

    public Dictionary<string, object?> GetContact(string entryId)
        => OutlookComInvoker.Run(() => GetContactCore(entryId));

    private Dictionary<string, object?> GetContactCore(string entryId)
    {
        var ns = GetNamespace();
        dynamic item;
        try
        {
            item = ns.GetItemFromID(entryId);
        }
        catch
        {
            OutlookComHost.Release(ns);
            throw new InvalidOperationException($"Contact not found with ID: {entryId}");
        }
        try
        {
            return ContactToDict(item);
        }
        finally
        {
            OutlookComHost.Release(item);
            OutlookComHost.Release(ns);
        }
    }

    public string CreateContact(string? firstName, string? lastName, string? email,
        string? phone, string? mobilePhone, string? company, string? jobTitle,
        string? businessAddress, string? notes, string? account = null)
        => OutlookComInvoker.Run(() => CreateContactCore(firstName, lastName, email, phone, mobilePhone, company, jobTitle, businessAddress, notes, account));

    private string CreateContactCore(string? firstName, string? lastName, string? email,
        string? phone, string? mobilePhone, string? company, string? jobTitle,
        string? businessAddress, string? notes, string? account)
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

        // Re-fetch after save: Outlook reassigns EntryID after first save to Exchange
        string tempId = (string)contact.EntryID;
        Marshal.ReleaseComObject(contact);
        var ns = GetNamespace();
        dynamic saved = ns.GetItemFromID(tempId);
        try
        {
            string stableId = (string)saved.EntryID;
            return stableId;
        }
        finally
        {
            Marshal.ReleaseComObject(saved);
            OutlookComHost.Release(ns);
        }
    }

    public bool UpdateContact(string entryId, string? firstName, string? lastName,
        string? email, string? phone, string? mobilePhone, string? company,
        string? jobTitle, string? businessAddress, string? notes)
        => OutlookComInvoker.Run(() => UpdateContactCore(entryId, firstName, lastName, email, phone, mobilePhone, company, jobTitle, businessAddress, notes));

    private bool UpdateContactCore(string entryId, string? firstName, string? lastName,
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
            OutlookComHost.Release(ns);
            throw new InvalidOperationException($"Contact not found with ID: {entryId}");
        }

        try
        {
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
            return true;
        }
        finally
        {
            Marshal.ReleaseComObject(contact);
            OutlookComHost.Release(ns);
        }
    }

    public bool DeleteContact(string entryId)
        => OutlookComInvoker.Run(() => DeleteContactCore(entryId));

    private bool DeleteContactCore(string entryId)
    {
        var ns = GetNamespace();
        dynamic contact;
        try
        {
            contact = ns.GetItemFromID(entryId);
        }
        catch
        {
            OutlookComHost.Release(ns);
            throw new InvalidOperationException($"Contact not found with ID: {entryId}");
        }

        try
        {
            contact.Delete();
        }
        catch (System.Runtime.InteropServices.COMException) { /* item deleted but COM reports a move error — ignore */ }
        finally
        {
            try { Marshal.ReleaseComObject(contact); } catch { }
            OutlookComHost.Release(ns);
        }
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

}
