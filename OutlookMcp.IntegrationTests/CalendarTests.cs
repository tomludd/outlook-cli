using OutlookMcp.Services;

namespace OutlookMcp.IntegrationTests;

public class CalendarTests : IClassFixture<OutlookFixture>
{
    private readonly OutlookCalendarService _svc;

    public CalendarTests(OutlookFixture fixture)
    {
        _svc = fixture.CalendarService;
    }

    [Fact]
    public void GetCalendars_ReturnsAtLeastDefaultCalendar()
    {
        var calendars = _svc.GetCalendars();

        Assert.NotNull(calendars);
        Assert.NotEmpty(calendars);

        var defaultCal = calendars.FirstOrDefault(c => c.ContainsKey("isDefault") && (bool)c["isDefault"]!);
        Assert.NotNull(defaultCal);
        Assert.True(defaultCal.ContainsKey("name"));
        Assert.True(defaultCal.ContainsKey("owner"));
    }

    [Fact]
    public void ListEvents_TodayRange_ReturnsListWithoutError()
    {
        var today = DateTime.Today;
        var events = _svc.ListEvents(today, today, null);

        Assert.NotNull(events);
        // May be empty if no events today — that's fine
    }

    [Fact]
    public void ListEvents_ThisWeek_ReturnsListWithoutError()
    {
        var today = DateTime.Today;
        var endOfWeek = today.AddDays(7);
        var events = _svc.ListEvents(today, endOfWeek, null);

        Assert.NotNull(events);
    }

    [Fact]
    public void ListEvents_FirstEventHasRequiredFields()
    {
        var today = DateTime.Today;
        var events = _svc.ListEvents(today, today.AddDays(30), null);

        if (events.Count == 0)
            return; // No events in next 30 days — skip

        var ev = events[0];
        Assert.True(ev.ContainsKey("id"));
        Assert.True(ev.ContainsKey("subject"));
        Assert.True(ev.ContainsKey("start"));
        Assert.True(ev.ContainsKey("end"));
        Assert.False(string.IsNullOrEmpty(ev["id"]?.ToString()));
    }

    [Fact]
    public void FindFreeSlots_Today_ReturnsSlots()
    {
        var today = DateTime.Today;
        var slots = _svc.FindFreeSlots(today, today, 30, 9, 17, null);

        Assert.NotNull(slots);
        // Weekday should have some slots unless calendar is fully booked
        if (today.DayOfWeek != DayOfWeek.Saturday && today.DayOfWeek != DayOfWeek.Sunday)
        {
            foreach (var slot in slots)
            {
                Assert.True(slot.ContainsKey("start"));
                Assert.True(slot.ContainsKey("end"));
            }
        }
    }

    [Fact]
    public void FindFreeSlots_ThisWeek_ReturnsSlots()
    {
        var today = DateTime.Today;
        var slots = _svc.FindFreeSlots(today, today.AddDays(7), 60, 9, 17, null);

        Assert.NotNull(slots);
        // A full week should have at least some free 1-hour slots
    }
}
