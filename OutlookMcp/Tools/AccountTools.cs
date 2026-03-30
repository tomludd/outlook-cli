using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;
using OutlookMcp.Services;

namespace OutlookMcp.Tools;

[McpServerToolType]
public class AccountTools
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    [McpServerTool(Name = "list_accounts"), Description("List all available Outlook accounts/stores. Use account display names with the 'account' parameter in other tools.")]
    public string ListAccounts()
    {
        using var svc = new OutlookCalendarService();
        var accounts = svc.ListAccounts();
        return JsonSerializer.Serialize(accounts, JsonOptions);
    }
}
