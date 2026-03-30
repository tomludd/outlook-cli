using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;
using OutlookMcp.Services;

namespace OutlookMcp.Tools;

[McpServerToolType]
public class ContactTools
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    [McpServerTool(Name = "list_contacts"), Description("List contacts from the Outlook contacts folder.")]
    public string ListContacts(
        [Description("Number of contacts to return (defaults to 50, max 500)")] int count = 50,
        [Description("Account name to use (optional, use list_accounts to see available accounts)")] string? account = null)
    {
        count = Math.Clamp(count, 1, 500);
        using var svc = new OutlookContactService();
        var contacts = svc.ListContacts(count, account);
        return JsonSerializer.Serialize(contacts, JsonOptions);
    }

    [McpServerTool(Name = "search_contacts"), Description("Search contacts by name, email, or company.")]
    public string SearchContacts(
        [Description("Search query")] string query,
        [Description("Maximum results to return (defaults to 20, max 100)")] int maxResults = 20,
        [Description("Account name to use (optional, use list_accounts to see available accounts)")] string? account = null)
    {
        maxResults = Math.Clamp(maxResults, 1, 100);
        using var svc = new OutlookContactService();
        var contacts = svc.SearchContacts(query, maxResults, account);
        return JsonSerializer.Serialize(contacts, JsonOptions);
    }

    [McpServerTool(Name = "get_contact"), Description("Get full details of a contact by its ID.")]
    public string GetContact(
        [Description("Contact ID (EntryID from list_contacts or search_contacts)")] string contactId)
    {
        using var svc = new OutlookContactService();
        var contact = svc.GetContact(contactId);
        return JsonSerializer.Serialize(contact, JsonOptions);
    }

    [McpServerTool(Name = "create_contact"), Description("Create a new contact in Outlook.")]
    public string CreateContact(
        [Description("First name (optional)")] string? firstName = null,
        [Description("Last name (optional)")] string? lastName = null,
        [Description("Email address (optional)")] string? email = null,
        [Description("Business phone number (optional)")] string? phone = null,
        [Description("Mobile phone number (optional)")] string? mobilePhone = null,
        [Description("Company name (optional)")] string? company = null,
        [Description("Job title (optional)")] string? jobTitle = null,
        [Description("Business address (optional)")] string? businessAddress = null,
        [Description("Notes (optional)")] string? notes = null,
        [Description("Account name to use (optional, use list_accounts to see available accounts)")] string? account = null)
    {
        using var svc = new OutlookContactService();
        var id = svc.CreateContact(firstName, lastName, email, phone, mobilePhone, company, jobTitle, businessAddress, notes, account);
        return JsonSerializer.Serialize(new { success = true, contactId = id }, JsonOptions);
    }

    [McpServerTool(Name = "update_contact"), Description("Update an existing contact. Only pass the fields you want to change.")]
    public string UpdateContact(
        [Description("Contact ID (EntryID)")] string contactId,
        [Description("First name (optional)")] string? firstName = null,
        [Description("Last name (optional)")] string? lastName = null,
        [Description("Email address (optional)")] string? email = null,
        [Description("Business phone number (optional)")] string? phone = null,
        [Description("Mobile phone number (optional)")] string? mobilePhone = null,
        [Description("Company name (optional)")] string? company = null,
        [Description("Job title (optional)")] string? jobTitle = null,
        [Description("Business address (optional)")] string? businessAddress = null,
        [Description("Notes (optional)")] string? notes = null)
    {
        using var svc = new OutlookContactService();
        var result = svc.UpdateContact(contactId, firstName, lastName, email, phone, mobilePhone, company, jobTitle, businessAddress, notes);
        return JsonSerializer.Serialize(new { success = result }, JsonOptions);
    }

    [McpServerTool(Name = "delete_contact"), Description("Delete a contact by its ID.")]
    public string DeleteContact(
        [Description("Contact ID (EntryID)")] string contactId)
    {
        using var svc = new OutlookContactService();
        var result = svc.DeleteContact(contactId);
        return JsonSerializer.Serialize(new { success = result }, JsonOptions);
    }
}
