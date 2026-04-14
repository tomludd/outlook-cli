using System.CommandLine;
using Outlook.COM;

namespace Outlook.Cli;

public static class AccountsCommand
{
    public static Command Build()
    {
        var cmd = new Command("accounts", "List available Outlook accounts");
        cmd.SetHandler(() =>
        {
            try
            {
                using var calService = new OutlookCalendarService();
                var accounts = calService.ListAccounts();
                if (accounts.Count == 0)
                {
                    Console.WriteLine("No Outlook accounts found.");
                    return;
                }

                foreach (var account in accounts)
                {
                    var name = account.TryGetValue("displayName", out var n) ? n?.ToString() : null;
                    if (!string.IsNullOrEmpty(name))
                        Console.WriteLine(name);
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Could not connect to Outlook: {ex.Message}");
            }
        });
        return cmd;
    }
}
