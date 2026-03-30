using OutlookMcp.Services;

namespace OutlookMcp.IntegrationTests;

public class AccountTests : IClassFixture<OutlookFixture>
{
    private readonly OutlookCalendarService _svc;
    private readonly ITestOutputHelper _output;

    public AccountTests(OutlookFixture fixture, ITestOutputHelper output)
    {
        _svc = fixture.CalendarService;
        _output = output;
    }

    [Fact]
    public void ListAccounts_ReturnsAtLeastOneAccount()
    {
        var accounts = _svc.ListAccounts();

        _output.WriteLine($"Found {accounts.Count} account(s):");
        foreach (var account in accounts)
        {
            var name = account["displayName"]?.ToString();
            var isDefault = account.ContainsKey("isDefault") && (bool)account["isDefault"]!;
            _output.WriteLine($"  [{(isDefault ? "DEFAULT" : "      ")}] {name}");
        }

        Assert.NotNull(accounts);
        Assert.NotEmpty(accounts);
    }

    [Fact]
    public void ListAccounts_EachAccountHasRequiredFields()
    {
        var accounts = _svc.ListAccounts();

        foreach (var account in accounts)
        {
            var name = account["displayName"]?.ToString() ?? "(unnamed)";
            _output.WriteLine($"Account: {name}");
            _output.WriteLine($"  displayName : {account.GetValueOrDefault("displayName")}");
            _output.WriteLine($"  storeId     : {account.GetValueOrDefault("storeId")?.ToString()?[..Math.Min(24, account.GetValueOrDefault("storeId")?.ToString()?.Length ?? 0)]}...");
            _output.WriteLine($"  isDefault   : {account.GetValueOrDefault("isDefault")}");

            Assert.True(account.ContainsKey("displayName"), $"Account missing 'displayName'");
            Assert.True(account.ContainsKey("storeId"), $"Account missing 'storeId'");
            Assert.True(account.ContainsKey("isDefault"), $"Account missing 'isDefault'");
            Assert.False(string.IsNullOrEmpty(account["displayName"]?.ToString()), "displayName should not be empty");
            Assert.False(string.IsNullOrEmpty(account["storeId"]?.ToString()), "storeId should not be empty");
        }
    }

    [Fact]
    public void ListAccounts_ExactlyOneDefaultAccount()
    {
        var accounts = _svc.ListAccounts();

        var defaults = accounts.Where(a => a.ContainsKey("isDefault") && (bool)a["isDefault"]!).ToList();

        _output.WriteLine($"Total accounts : {accounts.Count}");
        _output.WriteLine($"Default accounts: {defaults.Count}");
        foreach (var d in defaults)
            _output.WriteLine($"  Default: {d["displayName"]}");

        Assert.Single(defaults);
    }

    [Fact]
    public void ListAccounts_LogAllDetails()
    {
        var accounts = _svc.ListAccounts();

        _output.WriteLine("=== All Outlook Accounts ===");
        for (int i = 0; i < accounts.Count; i++)
        {
            var account = accounts[i];
            _output.WriteLine($"[{i + 1}] {account["displayName"]}");
            foreach (var kv in account)
            {
                var value = kv.Key == "storeId"
                    ? kv.Value?.ToString()?[..Math.Min(32, kv.Value?.ToString()?.Length ?? 0)] + "..."
                    : kv.Value?.ToString();
                _output.WriteLine($"    {kv.Key}: {value}");
            }
        }

        Assert.NotEmpty(accounts);
    }
}
