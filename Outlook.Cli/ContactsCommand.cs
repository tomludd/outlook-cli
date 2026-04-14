using System.CommandLine;
using System.Text.Json;
using Outlook.COM;

namespace Outlook.Cli;

public static class ContactsCommand
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    public static Command Build()
    {
        var cmd = new Command("contacts", "Manage Outlook contacts");
        cmd.Subcommands.Add(BuildList());
        cmd.Subcommands.Add(BuildSearch());
        cmd.Subcommands.Add(BuildGet());
        cmd.Subcommands.Add(BuildCreate());
        cmd.Subcommands.Add(BuildUpdate());
        cmd.Subcommands.Add(BuildDelete());
        return cmd;
    }

    private static Command BuildList()
    {
        var countOpt   = new Option<int>("--count",   "Number of contacts (max 500)") { DefaultValueFactory = _ => 50 };
        var accountOpt = new Option<string?>("--account", "Account display name (omit for all)");

        var cmd = new Command("list", "List contacts");
        cmd.Options.Add(countOpt); cmd.Options.Add(accountOpt);
        cmd.SetAction(ctx =>
        {
            var count   = Math.Clamp(ctx.GetValue(countOpt), 1, 500);
            var account = ctx.GetValue(accountOpt);
            using var svc = new OutlookContactService();
            var contacts = svc.ListContacts(count, account);
            Console.WriteLine(JsonSerializer.Serialize(contacts, JsonOptions));
        });
        return cmd;
    }

    private static Command BuildSearch()
    {
        var queryArg   = new Argument<string>("query") { Description = "Search by name, email, or company" };
        var maxOpt     = new Option<int>("--max",      "Maximum results (max 100)") { DefaultValueFactory = _ => 20 };
        var accountOpt = new Option<string?>("--account", "Account display name (omit for all)");

        var cmd = new Command("search", "Search contacts");
        cmd.Arguments.Add(queryArg);
        cmd.Options.Add(maxOpt); cmd.Options.Add(accountOpt);
        cmd.SetAction(ctx =>
        {
            var query   = ctx.GetValue(queryArg)!;
            var max     = Math.Clamp(ctx.GetValue(maxOpt), 1, 100);
            var account = ctx.GetValue(accountOpt);
            using var svc = new OutlookContactService();
            var contacts = svc.SearchContacts(query, max, account);
            Console.WriteLine(JsonSerializer.Serialize(contacts, JsonOptions));
        });
        return cmd;
    }

    private static Command BuildGet()
    {
        var idArg = new Argument<string>("id") { Description = "Contact ID (EntryID from list or search)" };
        var cmd   = new Command("get", "Get full contact details");
        cmd.Arguments.Add(idArg);
        cmd.SetAction(ctx =>
        {
            var id = ctx.GetValue(idArg)!;
            using var svc = new OutlookContactService();
            var contact = svc.GetContact(id);
            Console.WriteLine(JsonSerializer.Serialize(contact, JsonOptions));
        });
        return cmd;
    }

    private static Command BuildCreate()
    {
        var firstOpt   = new Option<string?>("--first",   "First name");
        var lastOpt    = new Option<string?>("--last",    "Last name");
        var emailOpt   = new Option<string?>("--email",   "Email address");
        var phoneOpt   = new Option<string?>("--phone",   "Business phone");
        var mobileOpt  = new Option<string?>("--mobile",  "Mobile phone");
        var companyOpt = new Option<string?>("--company", "Company name");
        var titleOpt   = new Option<string?>("--title",   "Job title");
        var addrOpt    = new Option<string?>("--address", "Business address");
        var notesOpt   = new Option<string?>("--notes",   "Notes");
        var accountOpt = new Option<string?>("--account", "Account display name");

        var cmd = new Command("create", "Create a new contact");
        cmd.Options.Add(firstOpt); cmd.Options.Add(lastOpt); cmd.Options.Add(emailOpt);
        cmd.Options.Add(phoneOpt); cmd.Options.Add(mobileOpt); cmd.Options.Add(companyOpt);
        cmd.Options.Add(titleOpt); cmd.Options.Add(addrOpt); cmd.Options.Add(notesOpt); cmd.Options.Add(accountOpt);
        cmd.SetAction(ctx =>
        {
            var first   = ctx.GetValue(firstOpt);
            var last    = ctx.GetValue(lastOpt);
            var email   = ctx.GetValue(emailOpt);
            var phone   = ctx.GetValue(phoneOpt);
            var mobile  = ctx.GetValue(mobileOpt);
            var company = ctx.GetValue(companyOpt);
            var title   = ctx.GetValue(titleOpt);
            var addr    = ctx.GetValue(addrOpt);
            var notes   = ctx.GetValue(notesOpt);
            var account = ctx.GetValue(accountOpt);
            using var svc = new OutlookContactService();
            var id = svc.CreateContact(first, last, email, phone, mobile, company, title, addr, notes, account);
            Console.WriteLine(JsonSerializer.Serialize(new { success = true, contactId = id }, JsonOptions));
        });
        return cmd;
    }

    private static Command BuildUpdate()
    {
        var idArg      = new Argument<string>("id") { Description = "Contact ID to update" };
        var firstOpt   = new Option<string?>("--first",   "First name");
        var lastOpt    = new Option<string?>("--last",    "Last name");
        var emailOpt   = new Option<string?>("--email",   "Email address");
        var phoneOpt   = new Option<string?>("--phone",   "Business phone");
        var mobileOpt  = new Option<string?>("--mobile",  "Mobile phone");
        var companyOpt = new Option<string?>("--company", "Company name");
        var titleOpt   = new Option<string?>("--title",   "Job title");
        var addrOpt    = new Option<string?>("--address", "Business address");
        var notesOpt   = new Option<string?>("--notes",   "Notes");

        var cmd = new Command("update", "Update an existing contact");
        cmd.Arguments.Add(idArg);
        cmd.Options.Add(firstOpt); cmd.Options.Add(lastOpt); cmd.Options.Add(emailOpt);
        cmd.Options.Add(phoneOpt); cmd.Options.Add(mobileOpt); cmd.Options.Add(companyOpt);
        cmd.Options.Add(titleOpt); cmd.Options.Add(addrOpt); cmd.Options.Add(notesOpt);
        cmd.SetAction(ctx =>
        {
            var id      = ctx.GetValue(idArg)!;
            var first   = ctx.GetValue(firstOpt);
            var last    = ctx.GetValue(lastOpt);
            var email   = ctx.GetValue(emailOpt);
            var phone   = ctx.GetValue(phoneOpt);
            var mobile  = ctx.GetValue(mobileOpt);
            var company = ctx.GetValue(companyOpt);
            var title   = ctx.GetValue(titleOpt);
            var addr    = ctx.GetValue(addrOpt);
            var notes   = ctx.GetValue(notesOpt);
            using var svc = new OutlookContactService();
            var result = svc.UpdateContact(id, first, last, email, phone, mobile, company, title, addr, notes);
            Console.WriteLine(JsonSerializer.Serialize(new { success = result }, JsonOptions));
        });
        return cmd;
    }

    private static Command BuildDelete()
    {
        var idArg = new Argument<string>("id") { Description = "Contact ID to delete" };
        var cmd   = new Command("delete", "Delete a contact");
        cmd.Arguments.Add(idArg);
        cmd.SetAction(ctx =>
        {
            var id = ctx.GetValue(idArg)!;
            using var svc = new OutlookContactService();
            var result = svc.DeleteContact(id);
            Console.WriteLine(JsonSerializer.Serialize(new { success = result }, JsonOptions));
        });
        return cmd;
    }
}

