using OutlookMcp.Services;

namespace OutlookMcp.IntegrationTests;

public class EmailTests : IClassFixture<OutlookFixture>
{
    private readonly OutlookMailService _svc;
    private readonly ITestOutputHelper _output;

    public EmailTests(OutlookFixture fixture, ITestOutputHelper output)
    {
        _svc = fixture.MailService;
        _output = output;
    }

    [Fact]
    public void ListEmails_Inbox_ReturnsListWithoutError()
    {
        var emails = _svc.ListEmails("inbox", 10, null, null);

        _output.WriteLine($"Inbox emails returned: {emails.Count}");
        foreach (var e in emails)
            _output.WriteLine($"  [{e.GetValueOrDefault("receivedTime")}] {e.GetValueOrDefault("from"),30}  {e.GetValueOrDefault("subject")}  [{e.GetValueOrDefault("account")}]");

        Assert.NotNull(emails);
    }

    [Fact]
    public void ListEmails_Sent_ReturnsListWithoutError()
    {
        var emails = _svc.ListEmails("sent", 10, null, null);

        _output.WriteLine($"Sent emails returned: {emails.Count}");
        foreach (var e in emails)
            _output.WriteLine($"  [{e.GetValueOrDefault("receivedTime")}] -> {e.GetValueOrDefault("to"),30}  {e.GetValueOrDefault("subject")}  [{e.GetValueOrDefault("account")}]");

        Assert.NotNull(emails);
    }

    [Fact]
    public void ListEmails_Drafts_ReturnsListWithoutError()
    {
        var emails = _svc.ListEmails("drafts", 10, null, null);

        Assert.NotNull(emails);
    }

    [Fact]
    public void ListEmails_FirstEmailHasRequiredFields()
    {
        var emails = _svc.ListEmails("inbox", 5, null, null);

        if (emails.Count == 0)
            return; // Empty inbox — skip

        var email = emails[0];
        Assert.True(email.ContainsKey("id"));
        Assert.True(email.ContainsKey("subject"));
        Assert.True(email.ContainsKey("from"));
        Assert.True(email.ContainsKey("receivedTime"));
        Assert.False(string.IsNullOrEmpty(email["id"]?.ToString()));
    }

    [Fact]
    public void GetEmail_WhenEmailExists_ReturnsBody()
    {
        var emails = _svc.ListEmails("inbox", 1, null, null);
        if (emails.Count == 0)
            return;

        var id = emails[0]["id"]!.ToString()!;
        var detail = _svc.GetEmail(id);

        Assert.NotNull(detail);
        Assert.Equal(id, detail["id"]?.ToString());
        Assert.True(detail.ContainsKey("body"), "Full email should include body");
    }

    [Fact]
    public void ListEmails_FilterBySubject_ReturnsFilteredResults()
    {
        // Get a subject from the first inbox email to use as filter
        var emails = _svc.ListEmails("inbox", 1, null, null);
        if (emails.Count == 0)
            return;

        var subject = emails[0]["subject"]?.ToString();
        if (string.IsNullOrEmpty(subject))
            return;

        // Use first word of subject as filter
        var filterWord = subject.Split(' ')[0];
        var filtered = _svc.ListEmails("inbox", 20, filterWord, null);

        Assert.NotNull(filtered);
        Assert.NotEmpty(filtered);
    }

    [Fact]
    public void SearchEmails_WithQuery_ReturnsListWithoutError()
    {
        var results = _svc.SearchEmails("meeting", 10);

        _output.WriteLine($"Search 'meeting' returned: {results.Count}");
        foreach (var e in results)
            _output.WriteLine($"  [{e.GetValueOrDefault("receivedTime")}] {e.GetValueOrDefault("subject")}  [{e.GetValueOrDefault("account")}]");

        Assert.NotNull(results);
    }
}
