using OutlookMcp.Services;

namespace OutlookMcp.IntegrationTests;

public class CalendarTests : IClassFixture<OutlookFixture>
{
    private readonly OutlookCalendarService _svc;
    private readonly ITestOutputHelper _output;

    public CalendarTests(OutlookFixture fixture, ITestOutputHelper output)
    {
        _svc = fixture.CalendarService;
        _output = output;
    }

    [Fact]
    public void GetCalendars_ReturnsAtLeastDefaultCalendar()
    {
        var calendars = _svc.GetCalendars();

        _output.WriteLine($"Found {calendars.Count} calendar(s):");
        foreach (var cal in calendars)
            _output.WriteLine($"  [{(cal.ContainsKey("isDefault") && (bool)cal["isDefault"]! ? "DEFAULT" : "      ")}] {cal.GetValueOrDefault("name")}");

        Assert.NotNull(calendars);
        Assert.NotEmpty(calendars);

        var defaultCal = calendars.FirstOrDefault(c => c.ContainsKey("isDefault") && (bool)c["isDefault"]!);
        Assert.NotNull(defaultCal);
        Assert.True(defaultCal.ContainsKey("name"));
        Assert.False(string.IsNullOrEmpty(defaultCal["name"]?.ToString()));
    }

    [Fact]
    public void ListEvents_TodayRange_ReturnsListWithoutError()
    {
        var today = DateTime.Today;
        var events = _svc.ListEvents(today, today, null);

        _output.WriteLine($"Events today ({today:yyyy-MM-dd}): {events.Count}");
        foreach (var ev in events)
            _output.WriteLine($"  {ev.GetValueOrDefault("start"),22}  {ev.GetValueOrDefault("subject")}  [{ev.GetValueOrDefault("account")}]");

        Assert.NotNull(events);
    }

    [Fact]
    public void ListEvents_ThisWeek_ReturnsListWithoutError()
    {
        var today = DateTime.Today;
        var endOfWeek = today.AddDays(7);
        var events = _svc.ListEvents(today, endOfWeek, null);

        _output.WriteLine($"Events this week ({today:yyyy-MM-dd} to {endOfWeek:yyyy-MM-dd}): {events.Count}");
        foreach (var ev in events)
            _output.WriteLine($"  {ev.GetValueOrDefault("start"),22}  {ev.GetValueOrDefault("subject")}  [{ev.GetValueOrDefault("account")}]");

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

        _output.WriteLine($"Free 30-min slots today: {slots.Count}");
        foreach (var slot in slots.Take(5))
            _output.WriteLine($"  {slot.GetValueOrDefault("start")} -> {slot.GetValueOrDefault("end")}");
        if (slots.Count > 5) _output.WriteLine($"  ... and {slots.Count - 5} more");

        Assert.NotNull(slots);
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

        _output.WriteLine($"Free 1-hour slots this week: {slots.Count}");
        foreach (var slot in slots.Take(5))
            _output.WriteLine($"  {slot.GetValueOrDefault("start")} -> {slot.GetValueOrDefault("end")}");
        if (slots.Count > 5) _output.WriteLine($"  ... and {slots.Count - 5} more");

        Assert.NotNull(slots);
    }
}
