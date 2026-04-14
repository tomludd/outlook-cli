using System.CommandLine;
using System.Text.Json;
using Outlook.COM;

namespace Outlook.Cli;

public static class EmailCommand
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    public static Command Build()
    {
        var cmd = new Command("email", "Manage Outlook emails");
        cmd.Subcommands.Add(BuildList());
        cmd.Subcommands.Add(BuildGet());
        cmd.Subcommands.Add(BuildSearch());
        cmd.Subcommands.Add(BuildSend());
        cmd.Subcommands.Add(BuildReply());
        cmd.Subcommands.Add(BuildForward());
        return cmd;
    }

    private static Command BuildList()
    {
        var folderOpt  = new Option<string?>("--folder") { Description = "Folder: inbox, sent, drafts, outbox (default: inbox)" };
        var countOpt   = new Option<int>("--count") { Description = "Number of emails (max 100)", DefaultValueFactory = _ => 20 };
        var subjectOpt = new Option<string?>("--subject") { Description = "Filter by subject" };
        var senderOpt  = new Option<string?>("--sender") { Description = "Filter by sender email" };
        var accountOpt = new Option<string?>("--account") { Description = "Account display name (omit for all accounts)" };
        var afterOpt   = new Option<string?>("--after") { Description = "Received on or after yyyy-MM-dd" };
        var beforeOpt  = new Option<string?>("--before") { Description = "Received before yyyy-MM-dd" };

        var cmd = new Command("list", "List recent emails");
        cmd.Options.Add(folderOpt); cmd.Options.Add(countOpt); cmd.Options.Add(subjectOpt);
        cmd.Options.Add(senderOpt); cmd.Options.Add(accountOpt); cmd.Options.Add(afterOpt); cmd.Options.Add(beforeOpt);
        cmd.SetAction(ctx =>
        {
            var folder  = ctx.GetValue(folderOpt);
            var count   = Math.Clamp(ctx.GetValue(countOpt), 1, 100);
            var subject = ctx.GetValue(subjectOpt);
            var sender  = ctx.GetValue(senderOpt);
            var account = ctx.GetValue(accountOpt);
            var after   = ctx.GetValue(afterOpt);
            var before  = ctx.GetValue(beforeOpt);
            using var svc = new OutlookMailService();
            var emails = svc.ListEmails(folder, count, subject, sender, account, after, before);
            Console.WriteLine(JsonSerializer.Serialize(emails, JsonOptions));
        });
        return cmd;
    }

    private static Command BuildGet()
    {
        var idArg = new Argument<string>("id") { Description = "Email ID (EntryID from list)" };
        var cmd   = new Command("get", "Get full email details including body");
        cmd.Arguments.Add(idArg);
        cmd.SetAction(ctx =>
        {
            var id = ctx.GetValue(idArg)!;
            using var svc = new OutlookMailService();
            var email = svc.GetEmail(id);
            Console.WriteLine(JsonSerializer.Serialize(email, JsonOptions));
        });
        return cmd;
    }

    private static Command BuildSearch()
    {
        var queryArg   = new Argument<string>("query") { Description = "Search keywords (subject, body, sender)" };
        var maxOpt     = new Option<int>("--max") { Description = "Maximum results (max 100)", DefaultValueFactory = _ => 20 };
        var accountOpt = new Option<string?>("--account") { Description = "Account display name (omit for all accounts)" };

        var cmd = new Command("search", "Search emails by keyword");
        cmd.Arguments.Add(queryArg);
        cmd.Options.Add(maxOpt); cmd.Options.Add(accountOpt);
        cmd.SetAction(ctx =>
        {
            var query   = ctx.GetValue(queryArg)!;
            var max     = Math.Clamp(ctx.GetValue(maxOpt), 1, 100);
            var account = ctx.GetValue(accountOpt);
            using var svc = new OutlookMailService();
            var emails = svc.SearchEmails(query, max, account);
            Console.WriteLine(JsonSerializer.Serialize(emails, JsonOptions));
        });
        return cmd;
    }

    private static Command BuildSend()
    {
        var toOpt          = new Option<string>("--to") { Description = "Recipient email(s), semicolon-separated", Required = true };
        var subjectOpt     = new Option<string>("--subject") { Description = "Email subject", Required = true };
        var bodyOpt        = new Option<string>("--body") { Description = "Email body", Required = true };
        var ccOpt          = new Option<string?>("--cc") { Description = "CC recipients, semicolon-separated" };
        var bccOpt         = new Option<string?>("--bcc") { Description = "BCC recipients, semicolon-separated" };
        var htmlOpt        = new Option<bool>("--html") { Description = "Body is HTML", DefaultValueFactory = _ => false };
        var importanceOpt  = new Option<string?>("--importance") { Description = "low, normal, high (default: normal)" };
        var attachmentsOpt = new Option<string?>("--attach") { Description = "File paths to attach, semicolon-separated" };
        var accountOpt     = new Option<string?>("--account") { Description = "Account display name to send from" };

        var cmd = new Command("send", "Send a new email");
        cmd.Options.Add(toOpt); cmd.Options.Add(subjectOpt); cmd.Options.Add(bodyOpt);
        cmd.Options.Add(ccOpt); cmd.Options.Add(bccOpt); cmd.Options.Add(htmlOpt);
        cmd.Options.Add(importanceOpt); cmd.Options.Add(attachmentsOpt); cmd.Options.Add(accountOpt);
        cmd.SetAction(ctx =>
        {
            var to          = ctx.GetValue(toOpt)!;
            var subject     = ctx.GetValue(subjectOpt)!;
            var body        = ctx.GetValue(bodyOpt)!;
            var cc          = ctx.GetValue(ccOpt);
            var bcc         = ctx.GetValue(bccOpt);
            var html        = ctx.GetValue(htmlOpt);
            var importance  = ctx.GetValue(importanceOpt);
            var attachments = ctx.GetValue(attachmentsOpt);
            var account     = ctx.GetValue(accountOpt);
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
        var idArg   = new Argument<string>("id") { Description = "Email ID to reply to" };
        var bodyOpt = new Option<string>("--body") { Description = "Reply body text", Required = true };
        var allOpt  = new Option<bool>("--all") { Description = "Reply to all recipients", DefaultValueFactory = _ => false };

        var cmd = new Command("reply", "Reply to an email");
        cmd.Arguments.Add(idArg);
        cmd.Options.Add(bodyOpt); cmd.Options.Add(allOpt);
        cmd.SetAction(ctx =>
        {
            var id   = ctx.GetValue(idArg)!;
            var body = ctx.GetValue(bodyOpt)!;
            var all  = ctx.GetValue(allOpt);
            using var svc = new OutlookMailService();
            svc.ReplyToEmail(id, body, all);
            Console.WriteLine(JsonSerializer.Serialize(new { success = true, message = all ? "Reply-all sent." : "Reply sent." }, JsonOptions));
        });
        return cmd;
    }

    private static Command BuildForward()
    {
        var idArg   = new Argument<string>("id") { Description = "Email ID to forward" };
        var toOpt   = new Option<string>("--to") { Description = "Recipient email(s), semicolon-separated", Required = true };
        var bodyOpt = new Option<string?>("--body") { Description = "Additional text to prepend" };

        var cmd = new Command("forward", "Forward an email");
        cmd.Arguments.Add(idArg);
        cmd.Options.Add(toOpt); cmd.Options.Add(bodyOpt);
        cmd.SetAction(ctx =>
        {
            var id   = ctx.GetValue(idArg)!;
            var to   = ctx.GetValue(toOpt)!;
            var body = ctx.GetValue(bodyOpt);
            using var svc = new OutlookMailService();
            svc.ForwardEmail(id, to, body);
            Console.WriteLine(JsonSerializer.Serialize(new { success = true, message = "Email forwarded." }, JsonOptions));
        });
        return cmd;
    }
}



