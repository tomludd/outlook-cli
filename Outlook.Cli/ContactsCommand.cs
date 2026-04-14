using System.CommandLine;
using System.CommandLine.Invocation;
using System.Text.Json;
using Outlook.COM;

namespace Outlook.Cli;

public static class ContactsCommand
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    public static Command Build()
    {
        var cmd = new Command("contacts", "Manage Outlook contacts");
        cmd.AddCommand(BuildList());
        cmd.AddCommand(BuildSearch());
        cmd.AddCommand(BuildGet());
        cmd.AddCommand(BuildCreate());
        cmd.AddCommand(BuildUpdate());
        cmd.AddCommand(BuildDelete());
        return cmd;
    }

    private static Command BuildList()
    {
        var countOpt   = new Option<int>("--count",   getDefaultValue: () => 50, description: "Number of contacts (max 500)");
        var accountOpt = new Option<string?>("--account", "Account display name (omit for all)");

        var cmd = new Command("list", "List contacts") { countOpt, accountOpt };
        cmd.SetHandler((int count, string? account) =>
        {
            count = Math.Clamp(count, 1, 500);
            using var svc = new OutlookContactService();
            var contacts = svc.ListContacts(count, account);
            Console.WriteLine(JsonSerializer.Serialize(contacts, JsonOptions));
        }, countOpt, accountOpt);
        return cmd;
    }

    private static Command BuildSearch()
    {
        var queryArg   = new Argument<string>("query", "Search by name, email, or company");
        var maxOpt     = new Option<int>("--max",     getDefaultValue: () => 20, description: "Maximum results (max 100)");
        var accountOpt = new Option<string?>("--account", "Account display name (omit for all)");

        var cmd = new Command("search", "Search contacts") { queryArg, maxOpt, accountOpt };
        cmd.SetHandler((string query, int max, string? account) =>
        {
            max = Math.Clamp(max, 1, 100);
            using var svc = new OutlookContactService();
            var contacts = svc.SearchContacts(query, max, account);
            Console.WriteLine(JsonSerializer.Serialize(contacts, JsonOptions));
        }, queryArg, maxOpt, accountOpt);
        return cmd;
    }

    private static Command BuildGet()
    {
        var idArg = new Argument<string>("id", "Contact ID (EntryID from list or search)");
        var cmd   = new Command("get", "Get full contact details") { idArg };
        cmd.SetHandler((string id) =>
        {
            using var svc = new OutlookContactService();
            var contact = svc.GetContact(id);
            Console.WriteLine(JsonSerializer.Serialize(contact, JsonOptions));
        }, idArg);
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

        var cmd = new Command("create", "Create a new contact") { firstOpt, lastOpt, emailOpt, phoneOpt, mobileOpt, companyOpt, titleOpt, addrOpt, notesOpt, accountOpt };
        cmd.SetHandler((InvocationContext ctx) =>
        {
            var first   = ctx.ParseResult.GetValueForOption(firstOpt);
            var last    = ctx.ParseResult.GetValueForOption(lastOpt);
            var email   = ctx.ParseResult.GetValueForOption(emailOpt);
            var phone   = ctx.ParseResult.GetValueForOption(phoneOpt);
            var mobile  = ctx.ParseResult.GetValueForOption(mobileOpt);
            var company = ctx.ParseResult.GetValueForOption(companyOpt);
            var title   = ctx.ParseResult.GetValueForOption(titleOpt);
            var addr    = ctx.ParseResult.GetValueForOption(addrOpt);
            var notes   = ctx.ParseResult.GetValueForOption(notesOpt);
            var account = ctx.ParseResult.GetValueForOption(accountOpt);
            using var svc = new OutlookContactService();
            var id = svc.CreateContact(first, last, email, phone, mobile, company, title, addr, notes, account);
            Console.WriteLine(JsonSerializer.Serialize(new { success = true, contactId = id }, JsonOptions));
        });
        return cmd;
    }

    private static Command BuildUpdate()
    {
        var idArg      = new Argument<string>("id", "Contact ID to update");
        var firstOpt   = new Option<string?>("--first",   "First name");
        var lastOpt    = new Option<string?>("--last",    "Last name");
        var emailOpt   = new Option<string?>("--email",   "Email address");
        var phoneOpt   = new Option<string?>("--phone",   "Business phone");
        var mobileOpt  = new Option<string?>("--mobile",  "Mobile phone");
        var companyOpt = new Option<string?>("--company", "Company name");
        var titleOpt   = new Option<string?>("--title",   "Job title");
        var addrOpt    = new Option<string?>("--address", "Business address");
        var notesOpt   = new Option<string?>("--notes",   "Notes");

        var cmd = new Command("update", "Update an existing contact") { idArg, firstOpt, lastOpt, emailOpt, phoneOpt, mobileOpt, companyOpt, titleOpt, addrOpt, notesOpt };
        cmd.SetHandler((InvocationContext ctx) =>
        {
            var id      = ctx.ParseResult.GetValueForArgument(idArg);
            var first   = ctx.ParseResult.GetValueForOption(firstOpt);
            var last    = ctx.ParseResult.GetValueForOption(lastOpt);
            var email   = ctx.ParseResult.GetValueForOption(emailOpt);
            var phone   = ctx.ParseResult.GetValueForOption(phoneOpt);
            var mobile  = ctx.ParseResult.GetValueForOption(mobileOpt);
            var company = ctx.ParseResult.GetValueForOption(companyOpt);
            var title   = ctx.ParseResult.GetValueForOption(titleOpt);
            var addr    = ctx.ParseResult.GetValueForOption(addrOpt);
            var notes   = ctx.ParseResult.GetValueForOption(notesOpt);
            using var svc = new OutlookContactService();
            var result = svc.UpdateContact(id, first, last, email, phone, mobile, company, title, addr, notes);
            Console.WriteLine(JsonSerializer.Serialize(new { success = result }, JsonOptions));
        });
        return cmd;
    }

    private static Command BuildDelete()
    {
        var idArg = new Argument<string>("id", "Contact ID to delete");
        var cmd   = new Command("delete", "Delete a contact") { idArg };
        cmd.SetHandler((string id) =>
        {
            using var svc = new OutlookContactService();
            var result = svc.DeleteContact(id);
            Console.WriteLine(JsonSerializer.Serialize(new { success = result }, JsonOptions));
        }, idArg);
        return cmd;
    }
}
