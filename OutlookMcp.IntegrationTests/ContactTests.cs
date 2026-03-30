using OutlookMcp.Services;

namespace OutlookMcp.IntegrationTests;

public class ContactTests : IClassFixture<OutlookFixture>
{
    private readonly OutlookContactService _svc;

    public ContactTests(OutlookFixture fixture)
    {
        _svc = fixture.ContactService;
    }

    [Fact]
    public void ListContacts_ReturnsListWithoutError()
    {
        var contacts = _svc.ListContacts(20);

        Assert.NotNull(contacts);
    }

    [Fact]
    public void ListContacts_FirstContactHasRequiredFields()
    {
        var contacts = _svc.ListContacts(5);

        if (contacts.Count == 0)
            return; // No contacts — skip

        var contact = contacts[0];
        Assert.True(contact.ContainsKey("id"));
        Assert.True(contact.ContainsKey("fullName"));
        Assert.False(string.IsNullOrEmpty(contact["id"]?.ToString()));
    }

    [Fact]
    public void GetContact_WhenContactExists_ReturnsDetails()
    {
        var contacts = _svc.ListContacts(1);
        if (contacts.Count == 0)
            return;

        var id = contacts[0]["id"]!.ToString()!;
        var detail = _svc.GetContact(id);

        Assert.NotNull(detail);
        Assert.Equal(id, detail["id"]?.ToString());
        Assert.True(detail.ContainsKey("fullName"));
        Assert.True(detail.ContainsKey("email"));
    }

    [Fact]
    public void SearchContacts_WithQuery_ReturnsListWithoutError()
    {
        // Get a name from the first contact to search for
        var contacts = _svc.ListContacts(1);
        if (contacts.Count == 0)
        {
            // Search for a generic term
            var results = _svc.SearchContacts("test", 10);
            Assert.NotNull(results);
            return;
        }

        var name = contacts[0]["fullName"]?.ToString();
        if (string.IsNullOrEmpty(name))
            return;

        var searchWord = name.Split(' ')[0];
        var found = _svc.SearchContacts(searchWord, 10);

        Assert.NotNull(found);
        Assert.NotEmpty(found);
    }

    [Fact]
    public void SearchContacts_NoMatch_ReturnsEmptyList()
    {
        var results = _svc.SearchContacts("zzz_nonexistent_xyzzy_12345", 10);

        Assert.NotNull(results);
        Assert.Empty(results);
    }
}
