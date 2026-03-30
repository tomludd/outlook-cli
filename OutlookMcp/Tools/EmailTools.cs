using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;
using OutlookMcp.Services;

namespace OutlookMcp.Tools;

[McpServerToolType]
public class EmailTools
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    [McpServerTool(Name = "list_emails"), Description("List recent emails from a mail folder. Returns newest first. Queries all accounts by default.")]
    public string ListEmails(
        [Description("Folder name: inbox, sent, drafts, or outbox (defaults to inbox)")] string? folder = null,
        [Description("Number of emails to return (defaults to 20, max 100)")] int count = 20,
        [Description("Filter by subject (optional)")] string? filterSubject = null,
        [Description("Filter by sender email (optional)")] string? filterSender = null,
        [Description("Account displayName to query (from list_accounts, e.g. 'tommy.kihlstrom@thon.no'). Omit to query all accounts.")] string? account = null)
    {
        count = Math.Clamp(count, 1, 100);
        using var svc = new OutlookMailService();
        var emails = svc.ListEmails(folder, count, filterSubject, filterSender, account);
        return JsonSerializer.Serialize(emails, JsonOptions);
    }

    [McpServerTool(Name = "get_email"), Description("Get the full details of an email by its ID, including the body.")]
    public string GetEmail(
        [Description("Email ID (EntryID from list_emails)")] string emailId)
    {
        using var svc = new OutlookMailService();
        var email = svc.GetEmail(emailId);
        return JsonSerializer.Serialize(email, JsonOptions);
    }

    [McpServerTool(Name = "search_emails"), Description("Search emails by keyword across subject, body, and sender. Searches all accounts by default.")]
    public string SearchEmails(
        [Description("Search query")] string query,
        [Description("Maximum results to return (defaults to 20, max 100)")] int maxResults = 20,
        [Description("Account displayName to query (from list_accounts, e.g. 'tommy.kihlstrom@thon.no'). Omit to query all accounts.")] string? account = null)
    {
        maxResults = Math.Clamp(maxResults, 1, 100);
        using var svc = new OutlookMailService();
        var emails = svc.SearchEmails(query, maxResults, account);
        return JsonSerializer.Serialize(emails, JsonOptions);
    }

    [McpServerTool(Name = "send_email"), Description("Send a new email.")]
    public string SendEmail(
        [Description("Recipient email addresses (semicolon-separated)")] string to,
        [Description("Email subject")] string subject,
        [Description("Email body text")] string body,
        [Description("CC recipients (semicolon-separated, optional)")] string? cc = null,
        [Description("BCC recipients (semicolon-separated, optional)")] string? bcc = null,
        [Description("Whether the body is HTML (defaults to false)")] bool isHtml = false,
        [Description("Importance: low, normal, high (defaults to normal)")] string? importance = null,
        [Description("File paths to attach (semicolon-separated, optional)")] string? attachments = null,
        [Description("Account displayName to send from (from list_accounts, e.g. 'tommy.kihlstrom@thon.no'). Omit to use the primary account.")] string? account = null)
    {
        var attachmentPaths = string.IsNullOrEmpty(attachments)
            ? null
            : attachments.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

        using var svc = new OutlookMailService();
        var id = svc.SendEmail(to, subject, body, cc, bcc, isHtml, importance, attachmentPaths, account);
        return JsonSerializer.Serialize(new { success = true, message = "Email sent successfully." }, JsonOptions);
    }

    [McpServerTool(Name = "reply_to_email"), Description("Reply to an existing email.")]
    public string ReplyToEmail(
        [Description("Email ID (EntryID) to reply to")] string emailId,
        [Description("Reply body text")] string body,
        [Description("Reply to all recipients (defaults to false)")] bool replyAll = false)
    {
        using var svc = new OutlookMailService();
        svc.ReplyToEmail(emailId, body, replyAll);
        return JsonSerializer.Serialize(new { success = true, message = replyAll ? "Reply-all sent." : "Reply sent." }, JsonOptions);
    }

    [McpServerTool(Name = "forward_email"), Description("Forward an existing email to new recipients.")]
    public string ForwardEmail(
        [Description("Email ID (EntryID) to forward")] string emailId,
        [Description("Recipient email addresses (semicolon-separated)")] string to,
        [Description("Additional body text to prepend (optional)")] string? body = null)
    {
        using var svc = new OutlookMailService();
        svc.ForwardEmail(emailId, to, body);
        return JsonSerializer.Serialize(new { success = true, message = "Email forwarded." }, JsonOptions);
    }
}
