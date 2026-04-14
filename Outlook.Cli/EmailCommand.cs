using System.CommandLine;
using System.CommandLine.Invocation;
using System.Text.Json;
using Outlook.COM;

namespace Outlook.Cli;

public static class EmailCommand
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    public static Command Build()
    {
        var cmd = new Command("email", "Manage Outlook emails");
        cmd.AddCommand(BuildList());
        cmd.AddCommand(BuildGet());
        cmd.AddCommand(BuildSearch());
        cmd.AddCommand(BuildSend());
        cmd.AddCommand(BuildReply());
        cmd.AddCommand(BuildForward());
        return cmd;
    }

    private static Command BuildList()
    {
        var folderOpt   = new Option<string?>("--folder",   "Folder: inbox, sent, drafts, outbox (default: inbox)");
        var countOpt    = new Option<int>("--count",        getDefaultValue: () => 20, description: "Number of emails (max 100)");
        var subjectOpt  = new Option<string?>("--subject",  "Filter by subject");
        var senderOpt   = new Option<string?>("--sender",   "Filter by sender email");
        var accountOpt  = new Option<string?>("--account",  "Account display name (omit for all accounts)");
        var afterOpt    = new Option<string?>("--after",    "Received on or after yyyy-MM-dd");
        var beforeOpt   = new Option<string?>("--before",   "Received before yyyy-MM-dd");

        var cmd = new Command("list", "List recent emails") { folderOpt, countOpt, subjectOpt, senderOpt, accountOpt, afterOpt, beforeOpt };
        cmd.SetHandler((string? folder, int count, string? subject, string? sender, string? account, string? after, string? before) =>
        {
            count = Math.Clamp(count, 1, 100);
            using var svc = new OutlookMailService();
            var emails = svc.ListEmails(folder, count, subject, sender, account, after, before);
            Console.WriteLine(JsonSerializer.Serialize(emails, JsonOptions));
        }, folderOpt, countOpt, subjectOpt, senderOpt, accountOpt, afterOpt, beforeOpt);
        return cmd;
    }

    private static Command BuildGet()
    {
        var idArg = new Argument<string>("id", "Email ID (EntryID from list)");
        var cmd = new Command("get", "Get full email details including body") { idArg };
        cmd.SetHandler((string id) =>
        {
            using var svc = new OutlookMailService();
            var email = svc.GetEmail(id);
            Console.WriteLine(JsonSerializer.Serialize(email, JsonOptions));
        }, idArg);
        return cmd;
    }

    private static Command BuildSearch()
    {
        var queryArg   = new Argument<string>("query", "Search keywords (subject, body, sender)");
        var maxOpt     = new Option<int>("--max",     getDefaultValue: () => 20, description: "Maximum results (max 100)");
        var accountOpt = new Option<string?>("--account", "Account display name (omit for all accounts)");

        var cmd = new Command("search", "Search emails by keyword") { queryArg, maxOpt, accountOpt };
        cmd.SetHandler((string query, int max, string? account) =>
        {
            max = Math.Clamp(max, 1, 100);
            using var svc = new OutlookMailService();
            var emails = svc.SearchEmails(query, max, account);
            Console.WriteLine(JsonSerializer.Serialize(emails, JsonOptions));
        }, queryArg, maxOpt, accountOpt);
        return cmd;
    }

    private static Command BuildSend()
    {
        var toOpt          = new Option<string>("--to",          "Recipient email(s), semicolon-separated") { IsRequired = true };
        var subjectOpt     = new Option<string>("--subject",     "Email subject") { IsRequired = true };
        var bodyOpt        = new Option<string>("--body",        "Email body") { IsRequired = true };
        var ccOpt          = new Option<string?>("--cc",         "CC recipients, semicolon-separated");
        var bccOpt         = new Option<string?>("--bcc",        "BCC recipients, semicolon-separated");
        var htmlOpt        = new Option<bool>("--html",          getDefaultValue: () => false, description: "Body is HTML");
        var importanceOpt  = new Option<string?>("--importance", "low, normal, high (default: normal)");
        var attachmentsOpt = new Option<string?>("--attach",     "File paths to attach, semicolon-separated");
        var accountOpt     = new Option<string?>("--account",    "Account display name to send from");

        var cmd = new Command("send", "Send a new email") { toOpt, subjectOpt, bodyOpt, ccOpt, bccOpt, htmlOpt, importanceOpt, attachmentsOpt, accountOpt };
        cmd.SetHandler((InvocationContext ctx) =>
        {
            var to          = ctx.ParseResult.GetValueForOption(toOpt)!;
            var subject     = ctx.ParseResult.GetValueForOption(subjectOpt)!;
            var body        = ctx.ParseResult.GetValueForOption(bodyOpt)!;
            var cc          = ctx.ParseResult.GetValueForOption(ccOpt);
            var bcc         = ctx.ParseResult.GetValueForOption(bccOpt);
            var html        = ctx.ParseResult.GetValueForOption(htmlOpt);
            var importance  = ctx.ParseResult.GetValueForOption(importanceOpt);
            var attachments = ctx.ParseResult.GetValueForOption(attachmentsOpt);
            var account     = ctx.ParseResult.GetValueForOption(accountOpt);
            var paths = string.IsNullOrEmpty(attachments)
                ? null
                : attachments.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            using var svc = new OutlookMailService();
            svc.SendEmail(to, subject, body, cc, bcc, html, importance, paths, account);
            Console.WriteLine(JsonSerializer.Serialize(new { success = true, message = "Email sent." }, JsonOptions));
        });
        return cmd;
    }

    private static Command BuildReply()
    {
        var idArg      = new Argument<string>("id", "Email ID to reply to");
        var bodyOpt    = new Option<string>("--body",     "Reply body text") { IsRequired = true };
        var allOpt     = new Option<bool>("--all",        getDefaultValue: () => false, description: "Reply to all recipients");

        var cmd = new Command("reply", "Reply to an email") { idArg, bodyOpt, allOpt };
        cmd.SetHandler((string id, string body, bool all) =>
        {
            using var svc = new OutlookMailService();
            svc.ReplyToEmail(id, body, all);
            Console.WriteLine(JsonSerializer.Serialize(new { success = true, message = all ? "Reply-all sent." : "Reply sent." }, JsonOptions));
        }, idArg, bodyOpt, allOpt);
        return cmd;
    }

    private static Command BuildForward()
    {
        var idArg   = new Argument<string>("id", "Email ID to forward");
        var toOpt   = new Option<string>("--to",   "Recipient email(s), semicolon-separated") { IsRequired = true };
        var bodyOpt = new Option<string?>("--body", "Additional text to prepend");

        var cmd = new Command("forward", "Forward an email") { idArg, toOpt, bodyOpt };
        cmd.SetHandler((string id, string to, string? body) =>
        {
            using var svc = new OutlookMailService();
            svc.ForwardEmail(id, to, body);
            Console.WriteLine(JsonSerializer.Serialize(new { success = true, message = "Email forwarded." }, JsonOptions));
        }, idArg, toOpt, bodyOpt);
        return cmd;
    }
}
