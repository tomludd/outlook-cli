using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace OutlookMcp.Services;

[SupportedOSPlatform("windows")]
public class OutlookMailService : IDisposable
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

    public List<Dictionary<string, object?>> ListEmails(string? folder, int count, string? filterSubject, string? filterSender, string? account = null)
    {
        var mailFolder = GetFolder(folder, account);
        var items = mailFolder.Items;
        items.Sort("[ReceivedTime]", true); // newest first

        // Apply DASL filter if provided
        if (!string.IsNullOrEmpty(filterSubject))
        {
            items = items.Restrict($"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{EscapeDasl(filterSubject)}%'");
        }
        else if (!string.IsNullOrEmpty(filterSender))
        {
            items = items.Restrict($"@SQL=\"urn:schemas:httpmail:fromemail\" LIKE '%{EscapeDasl(filterSender)}%'");
        }

        var emails = new List<Dictionary<string, object?>>();
        int limit = Math.Min(count, items.Count);
        for (int i = 1; i <= limit; i++)
        {
            var item = items.Item(i);
            try
            {
                emails.Add(MailToDict(item, includeBody: false));
            }
            catch
            {
                // Skip non-mail items (meeting requests, etc.)
            }
        }
        return emails;
    }

    public Dictionary<string, object?> GetEmail(string entryId)
    {
        var ns = GetNamespace();
        dynamic item;
        try
        {
            item = ns.GetItemFromID(entryId);
        }
        catch
        {
            throw new InvalidOperationException($"Email not found with ID: {entryId}");
        }
        return MailToDict(item, includeBody: true);
    }

    public string SendEmail(string to, string subject, string body, string? cc, string? bcc,
        bool isHtml, string? importance, string[]? attachmentPaths, string? account = null)
    {
        var app = GetOutlookApp();
        var mail = app.CreateItem(OlMailItem);

        // Set the sending account if specified
        if (!string.IsNullOrEmpty(account))
        {
            var accounts = app.Session.Accounts;
            bool found = false;
            for (int i = 1; i <= accounts.Count; i++)
            {
                var acc = accounts.Item(i);
                if (string.Equals((string)acc.DisplayName, account, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals((string)acc.SmtpAddress, account, StringComparison.OrdinalIgnoreCase))
                {
                    mail.SendUsingAccount = acc;
                    found = true;
                    break;
                }
            }
            if (!found)
                throw new InvalidOperationException($"Account not found: {account}. Use list_accounts to see available accounts.");
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
        Marshal.ReleaseComObject(mail);
        return entryId;
    }

    public string ReplyToEmail(string entryId, string body, bool replyAll)
    {
        var ns = GetNamespace();
        dynamic original;
        try
        {
            original = ns.GetItemFromID(entryId);
        }
        catch
        {
            throw new InvalidOperationException($"Email not found with ID: {entryId}");
        }

        var reply = replyAll ? original.ReplyAll() : original.Reply();
        reply.Body = body + reply.Body;
        reply.Send();

        string replyId = reply.EntryID ?? "";
        Marshal.ReleaseComObject(reply);
        Marshal.ReleaseComObject(original);
        return replyId;
    }

    public string ForwardEmail(string entryId, string to, string? additionalBody)
    {
        var ns = GetNamespace();
        dynamic original;
        try
        {
            original = ns.GetItemFromID(entryId);
        }
        catch
        {
            throw new InvalidOperationException($"Email not found with ID: {entryId}");
        }

        var fwd = original.Forward();
        fwd.To = to;
        if (!string.IsNullOrEmpty(additionalBody))
            fwd.Body = additionalBody + fwd.Body;

        fwd.Send();

        string fwdId = fwd.EntryID ?? "";
        Marshal.ReleaseComObject(fwd);
        Marshal.ReleaseComObject(original);
        return fwdId;
    }

    public List<Dictionary<string, object?>> SearchEmails(string query, int maxResults, string? account = null)
    {
        var ns = GetNamespace();
        var inbox = string.IsNullOrEmpty(account)
            ? ns.GetDefaultFolder(OlFolderInbox)
            : GetStoreFolder(account, OlFolderInbox);

        // Search across subject, body, and sender
        var filter = $"@SQL=(\"urn:schemas:httpmail:subject\" LIKE '%{EscapeDasl(query)}%' " +
                     $"OR \"urn:schemas:httpmail:textdescription\" LIKE '%{EscapeDasl(query)}%' " +
                     $"OR \"urn:schemas:httpmail:fromemail\" LIKE '%{EscapeDasl(query)}%')";

        var items = inbox.Items.Restrict(filter);
        items.Sort("[ReceivedTime]", true);

        var emails = new List<Dictionary<string, object?>>();
        int limit = Math.Min(maxResults, items.Count);
        for (int i = 1; i <= limit; i++)
        {
            try
            {
                emails.Add(MailToDict(items.Item(i), includeBody: false));
            }
            catch
            {
                // Skip non-mail items
            }
        }
        return emails;
    }

    private static Dictionary<string, object?> MailToDict(dynamic mail, bool includeBody)
    {
        var dict = new Dictionary<string, object?>
        {
            ["id"] = (string)mail.EntryID,
            ["subject"] = (string)mail.Subject,
            ["from"] = (string)mail.SenderEmailAddress,
            ["senderName"] = (string)mail.SenderName,
            ["to"] = (string)mail.To,
            ["cc"] = (string)mail.CC,
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
        for (int i = 1; i <= atts.Count; i++)
            attachments.Add((string)atts.Item(i).FileName);
        dict["attachments"] = attachments;

        if (includeBody)
            dict["body"] = (string)mail.Body;

        return dict;
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
