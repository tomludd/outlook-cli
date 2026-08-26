using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace Outlook.COM;

[SupportedOSPlatform("windows")]
public class OutlookMailService
{
    // Outlook folder constants
    private const int OlFolderInbox = 6;
    private const int OlFolderSentMail = 5;
    private const int OlFolderDrafts = 16;
    private const int OlFolderOutbox = 4;
    private const int OlMailItem = 0;
    private const int OlImportanceLow = 0;
    private const int OlImportanceNormal = 1;
    private const int OlImportanceHigh = 2;
    private const int OlByValue = 1;

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

    private dynamic GetFolder(string? folderName, string? account)
    {
        if (string.IsNullOrEmpty(account))
        {
            var ns = GetNamespace();
            return (folderName?.ToLowerInvariant()) switch
            {
                null or "" or "inbox" => ns.GetDefaultFolder(OlFolderInbox),
                "sent" or "sentmail" => ns.GetDefaultFolder(OlFolderSentMail),
                "drafts" => ns.GetDefaultFolder(OlFolderDrafts),
                "outbox" => ns.GetDefaultFolder(OlFolderOutbox),
                _ => throw new InvalidOperationException($"Unknown folder: {folderName}. Use inbox, sent, drafts, or outbox.")
            };
        }

        return (folderName?.ToLowerInvariant()) switch
        {
            null or "" or "inbox" => GetStoreFolder(account, OlFolderInbox),
            "sent" or "sentmail" => GetStoreFolder(account, OlFolderSentMail),
            "drafts" => GetStoreFolder(account, OlFolderDrafts),
            "outbox" => GetStoreFolder(account, OlFolderOutbox),
            _ => throw new InvalidOperationException($"Unknown folder: {folderName}. Use inbox, sent, drafts, or outbox.")
        };
    }

    public List<Dictionary<string, object?>> ListEmails(string? folder, int count, string? filterSubject, string? filterSender, string? account = null, string? receivedAfter = null, string? receivedBefore = null, bool includeBody = false)
        => OutlookComInvoker.Run(() => ListEmailsCore(folder, count, filterSubject, filterSender, account, receivedAfter, receivedBefore, includeBody));

    private List<Dictionary<string, object?>> ListEmailsCore(string? folder, int count, string? filterSubject, string? filterSender, string? account, string? receivedAfter, string? receivedBefore, bool includeBody)
    {
        if (!string.IsNullOrEmpty(account))
            return CollectEmails(GetFolder(folder, account), count, filterSubject, filterSender, account, receivedAfter, receivedBefore, includeBody);

        // Aggregate across all accounts when no account is specified
        int folderType = (folder?.ToLowerInvariant()) switch
        {
            "sent" or "sentmail" => OlFolderSentMail,
            "drafts" => OlFolderDrafts,
            "outbox" => OlFolderOutbox,
            _ => OlFolderInbox
        };

        var all = new List<Dictionary<string, object?>>();
        var ns = GetNamespace();
        var stores = ns.Stores;
        try
        {
            for (int i = 1; i <= stores.Count; i++)
            {
                dynamic? store = null;
                try { store = stores.Item(i); all.AddRange(CollectEmails(store.GetDefaultFolder(folderType), count, filterSubject, filterSender, (string)store.DisplayName, receivedAfter, receivedBefore, includeBody)); }
                catch { /* Store may not have this folder */ }
                finally { OutlookComHost.Release(store); }
            }
        }
        finally
        {
            OutlookComHost.Release(stores);
            OutlookComHost.Release(ns);
        }

        all.Sort((a, b) => string.Compare(b["receivedTime"]?.ToString(), a["receivedTime"]?.ToString(), StringComparison.Ordinal));
        return all.Take(count).ToList();
    }

    private List<Dictionary<string, object?>> CollectEmails(dynamic mailFolder, int count, string? filterSubject, string? filterSender, string? accountName, string? receivedAfter = null, string? receivedBefore = null, bool includeBody = false)
    {
        var items = mailFolder.Items;
        items.Sort("[ReceivedTime]", true); // newest first

        // Build combined DASL filter
        var conditions = new List<string>();
        if (!string.IsNullOrEmpty(filterSubject))
            conditions.Add($"\"urn:schemas:httpmail:subject\" LIKE '%{EscapeDasl(filterSubject)}%'");
        if (!string.IsNullOrEmpty(filterSender))
            conditions.Add($"\"urn:schemas:httpmail:fromemail\" LIKE '%{EscapeDasl(filterSender)}%'");
        if (!string.IsNullOrEmpty(receivedAfter) && DateTime.TryParse(receivedAfter, out var afterDate))
            conditions.Add($"\"urn:schemas:httpmail:datereceived\" >= '{afterDate:yyyy-MM-dd HH:mm}'");
        if (!string.IsNullOrEmpty(receivedBefore) && DateTime.TryParse(receivedBefore, out var beforeDate))
            conditions.Add($"\"urn:schemas:httpmail:datereceived\" < '{beforeDate:yyyy-MM-dd HH:mm}'");

        if (conditions.Count > 0)
            items = items.Restrict($"@SQL={string.Join(" AND ", conditions)}");

        var emails = new List<Dictionary<string, object?>>();
        int limit = Math.Min(count, items.Count);
        try
        {
            for (int i = 1; i <= limit; i++)
            {
                dynamic? item = null;
                try
                {
                    item = items.Item(i);
                    var email = MailToDict(item, includeBody);
                    if (accountName != null) email["account"] = accountName;
                    emails.Add(email);
                }
                catch { /* Skip non-mail items (meeting requests, etc.) */ }
                finally { OutlookComHost.Release(item); }
            }
        }
        finally { OutlookComHost.Release(items); }
        return emails;
    }

    public Dictionary<string, object?> GetEmail(string entryId)
        => OutlookComInvoker.Run(() => GetEmailCore(entryId));

    private Dictionary<string, object?> GetEmailCore(string entryId)
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
            throw new InvalidOperationException($"Email not found with ID: {entryId}");
        }
        try
        {
            return MailToDict(item, includeBody: true);
        }
        finally
        {
            OutlookComHost.Release(item);
            OutlookComHost.Release(ns);
        }
    }

    public string SendEmail(string to, string subject, string body, string? cc, string? bcc,
        bool isHtml, string? importance, string[]? attachmentPaths, string? account = null)
        => OutlookComInvoker.Run(() => SendEmailCore(to, subject, body, cc, bcc, isHtml, importance, attachmentPaths, account));

    private string SendEmailCore(string to, string subject, string body, string? cc, string? bcc,
        bool isHtml, string? importance, string[]? attachmentPaths, string? account)
    {
        var app = GetOutlookApp();
        var mail = app.CreateItem(OlMailItem);

        try
        {
            // Set the sending account if specified
            if (!string.IsNullOrEmpty(account))
            {
                var accounts = app.Session.Accounts;
                try
                {
                    bool found = false;
                    for (int i = 1; i <= accounts.Count; i++)
                    {
                        var acc = accounts.Item(i);
                        try
                        {
                            if (string.Equals((string)acc.DisplayName, account, StringComparison.OrdinalIgnoreCase) ||
                                string.Equals((string)acc.SmtpAddress, account, StringComparison.OrdinalIgnoreCase))
                            {
                                mail.SendUsingAccount = acc;
                                found = true;
                                break;
                            }
                        }
                        finally { OutlookComHost.Release(acc); }
                    }
                    if (!found)
                        throw new InvalidOperationException($"Account not found: {account}. Use list_accounts to see available accounts.");
                }
                finally { OutlookComHost.Release(accounts); }
            }

            mail.To = to;
            mail.Subject = subject;

            if (isHtml)
                mail.HTMLBody = body;
            else
                mail.Body = body;

            if (!string.IsNullOrEmpty(cc)) mail.CC = cc;
            if (!string.IsNullOrEmpty(bcc)) mail.BCC = bcc;

            mail.Importance = importance?.ToLowerInvariant() switch
            {
                "high" => OlImportanceHigh,
                "low" => OlImportanceLow,
                _ => OlImportanceNormal
            };

            if (attachmentPaths != null)
            {
                foreach (var path in attachmentPaths)
                {
                    if (!File.Exists(path))
                        throw new FileNotFoundException($"Attachment not found: {path}");
                    mail.Attachments.Add(path, OlByValue);
                }
            }

            mail.Send();
            string entryId = mail.EntryID ?? "";
            return entryId;
        }
        finally
        {
            Marshal.ReleaseComObject(mail);
        }
    }

    public string ReplyToEmail(string entryId, string body, bool replyAll)
        => OutlookComInvoker.Run(() => ReplyToEmailCore(entryId, body, replyAll));

    private string ReplyToEmailCore(string entryId, string body, bool replyAll)
    {
        var ns = GetNamespace();
        dynamic original;
        try
        {
            original = ns.GetItemFromID(entryId);
        }
        catch
        {
            OutlookComHost.Release(ns);
            throw new InvalidOperationException($"Email not found with ID: {entryId}");
        }

        dynamic? reply = null;
        try
        {
            reply = replyAll ? original.ReplyAll() : original.Reply();
            reply.Body = body + reply.Body;
            reply.Send();
            string replyId = reply.EntryID ?? "";
            return replyId;
        }
        finally
        {
            OutlookComHost.Release(reply);
            Marshal.ReleaseComObject(original);
            OutlookComHost.Release(ns);
        }
    }

    public string ForwardEmail(string entryId, string to, string? additionalBody)
        => OutlookComInvoker.Run(() => ForwardEmailCore(entryId, to, additionalBody));

    private string ForwardEmailCore(string entryId, string to, string? additionalBody)
    {
        var ns = GetNamespace();
        dynamic original;
        try
        {
            original = ns.GetItemFromID(entryId);
        }
        catch
        {
            OutlookComHost.Release(ns);
            throw new InvalidOperationException($"Email not found with ID: {entryId}");
        }

        dynamic? fwd = null;
        try
        {
            fwd = original.Forward();
            fwd.To = to;
            if (!string.IsNullOrEmpty(additionalBody))
                fwd.Body = additionalBody + fwd.Body;

            fwd.Send();
            string fwdId = fwd.EntryID ?? "";
            return fwdId;
        }
        finally
        {
            OutlookComHost.Release(fwd);
            Marshal.ReleaseComObject(original);
            OutlookComHost.Release(ns);
        }
    }

    public List<Dictionary<string, object?>> SearchEmails(string query, int maxResults, string? account = null, bool includeBody = false)
        => OutlookComInvoker.Run(() => SearchEmailsCore(query, maxResults, account, includeBody));

    private List<Dictionary<string, object?>> SearchEmailsCore(string query, int maxResults, string? account, bool includeBody)
    {
        var filter = $"@SQL=(\"urn:schemas:httpmail:subject\" LIKE '%{EscapeDasl(query)}%' " +
                     $"OR \"urn:schemas:httpmail:textdescription\" LIKE '%{EscapeDasl(query)}%' " +
                     $"OR \"urn:schemas:httpmail:fromemail\" LIKE '%{EscapeDasl(query)}%')";

        List<Dictionary<string, object?>> SearchFolder(dynamic inbox, string? accountName)
        {
            var items = inbox.Items.Restrict(filter);
            items.Sort("[ReceivedTime]", true);
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
                        var email = MailToDict(item, includeBody);
                        if (accountName != null) email["account"] = accountName;
                        results.Add(email);
                    }
                    catch { /* Skip non-mail items */ }
                    finally { OutlookComHost.Release(item); }
                }
            }
            finally { OutlookComHost.Release(items); }
            return results;
        }

        if (!string.IsNullOrEmpty(account))
            return SearchFolder(GetStoreFolder(account, OlFolderInbox), account);

        // Search across all accounts
        var ns = GetNamespace();
        var all = new List<Dictionary<string, object?>>();
        var stores = ns.Stores;
        try
        {
            for (int i = 1; i <= stores.Count; i++)
            {
                dynamic? store = null;
                try { store = stores.Item(i); all.AddRange(SearchFolder(store.GetDefaultFolder(OlFolderInbox), (string)store.DisplayName)); }
                catch { /* Store may not have inbox */ }
                finally { OutlookComHost.Release(store); }
            }
        }
        finally
        {
            OutlookComHost.Release(stores);
            OutlookComHost.Release(ns);
        }

        all.Sort((a, b) => string.Compare(b["receivedTime"]?.ToString(), a["receivedTime"]?.ToString(), StringComparison.Ordinal));
        return all.Take(maxResults).ToList();
    }

    private static Dictionary<string, object?> MailToDict(dynamic mail, bool includeBody)
    {
        var to = (string)mail.To;
        var cc = (string)mail.CC;

        var dict = new Dictionary<string, object?>
        {
            ["id"] = (string)mail.EntryID,
            ["subject"] = (string)mail.Subject,
            ["from"] = (string)mail.SenderEmailAddress,
            ["senderName"] = (string)mail.SenderName,
            ["to"] = includeBody ? to : (to?.Length > 80 ? to[..80] + "..." : to),
            ["cc"] = includeBody ? cc : (cc?.Length > 80 ? cc[..80] + "..." : cc),
            ["receivedTime"] = ((DateTime)mail.ReceivedTime).ToString("yyyy-MM-dd HH:mm"),
            ["isRead"] = (bool)mail.UnRead == false,
        };

        dict["importance"] = (int)mail.Importance switch
        {
            OlImportanceHigh => "High",
            OlImportanceLow => "Low",
            _ => "Normal"
        };

        // Attachments summary
        var attachments = new List<string>();
        var atts = mail.Attachments;
        try
        {
            for (int i = 1; i <= atts.Count; i++)
            {
                dynamic? att = null;
                try { att = atts.Item(i); attachments.Add((string)att.FileName); }
                finally { OutlookComHost.Release(att); }
            }
        }
        finally { OutlookComHost.Release(atts); }
        dict["attachments"] = attachments;

        if (includeBody)
            dict["body"] = (string)mail.Body;

        return dict;
    }

    private static string EscapeDasl(string value) =>
        value.Replace("'", "''").Replace("\"", "\"\"");

}
