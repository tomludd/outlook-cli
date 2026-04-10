using System.ComponentModel;
using System.Globalization;
using System.Text.Json;
using ModelContextProtocol.Server;
using Outlook.COM;

namespace Outlook.MCP;

[McpServerToolType]
public class CalendarTools
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    [McpServerTool(Name = "list_events"), Description("List calendar events within a specified date range. Returns subject, time, location, organizer, and meeting status. Use get_event for full body and attendees. Queries all accounts by default.")]
    public string ListEvents(
        [Description("Start date (yyyy-MM-dd)")] string startDate,
        [Description("End date (yyyy-MM-dd), use the same day as startDate to capture the full day")] string endDate,
        [Description("Account displayName to query (from list_accounts, e.g. 'tommy.kihlstrom@thon.no'). Omit to query all accounts.")] string? account = null)
    {
        var start = ParseDate(startDate);
        var end = ParseDate(endDate);
        if (end < start)
            return Error("End date cannot be before start date.");

        using var svc = new OutlookCalendarService();
        var events = svc.ListEvents(start, end, account);
        return JsonSerializer.Serialize(events, JsonOptions);
    }

    [McpServerTool(Name = "get_event"), Description("Get the full details of a calendar event by its ID, including body and attendees.")]
    public string GetEvent(
        [Description("Event ID (EntryID from list_events)")] string eventId)
    {
        using var svc = new OutlookCalendarService();
        var ev = svc.GetEvent(eventId);
        return JsonSerializer.Serialize(ev, JsonOptions);
    }

    [McpServerTool(Name = "create_event"), Description("Create a new calendar event or meeting.")]
    public string CreateEvent(
        [Description("Event subject/title")] string subject,
        [Description("Start date (yyyy-MM-dd)")] string startDate,
        [Description("Start time (HH:mm, 24-hour format)")] string startTime,
        [Description("End date (yyyy-MM-dd, optional, defaults to start date)")] string? endDate = null,
        [Description("End time (HH:mm, 24-hour format, optional, defaults to 30 min after start)")] string? endTime = null,
        [Description("Event location (optional)")] string? location = null,
        [Description("Event description/body (optional)")] string? body = null,
        [Description("Whether this is a meeting with attendees (optional)")] bool isMeeting = false,
        [Description("Semicolon-separated list of attendee email addresses (optional)")] string? attendees = null,
        [Description("Account displayName to create in (from list_accounts, e.g. 'tommy.kihlstrom@thon.no'). Omit to use the primary account.")] string? account = null)
    {
        var startDt = ParseDateTime(startDate, startTime);
        var endDt = !string.IsNullOrEmpty(endDate) && !string.IsNullOrEmpty(endTime)
            ? ParseDateTime(endDate, endTime)
            : !string.IsNullOrEmpty(endTime)
                ? ParseDateTime(startDate, endTime)
                : startDt.AddMinutes(30);

        if (endDt <= startDt)
            return Error("End time must be after start time.");

        using var svc = new OutlookCalendarService();
        var eventId = svc.CreateEvent(subject, startDt, endDt, location, body, isMeeting, attendees, account);
        return JsonSerializer.Serialize(new { success = true, eventId }, JsonOptions);
    }

    [McpServerTool(Name = "update_event"), Description("Update an existing calendar event. Only pass the fields you want to change.")]
    public string UpdateEvent(
        [Description("Event ID (EntryID from list_events)")] string eventId,
        [Description("New subject/title (optional)")] string? subject = null,
        [Description("New start date (yyyy-MM-dd, optional)")] string? startDate = null,
        [Description("New start time (HH:mm, 24-hour format, optional)")] string? startTime = null,
        [Description("New end date (yyyy-MM-dd, optional)")] string? endDate = null,
        [Description("New end time (HH:mm, 24-hour format, optional)")] string? endTime = null,
        [Description("New location (optional)")] string? location = null,
        [Description("New description/body (optional)")] string? body = null,
        [Description("Account displayName (from list_accounts). Usually not needed when an event ID is provided.")] string? account = null)
    {
        DateTime? startDt = null;
        DateTime? endDt = null;

        if (!string.IsNullOrEmpty(startDate) && !string.IsNullOrEmpty(startTime))
            startDt = ParseDateTime(startDate, startTime);
        else if (!string.IsNullOrEmpty(startDate))
            startDt = ParseDate(startDate); // date-only, service will preserve existing time

        if (!string.IsNullOrEmpty(endDate) && !string.IsNullOrEmpty(endTime))
            endDt = ParseDateTime(endDate, endTime);
        else if (!string.IsNullOrEmpty(endDate))
            endDt = ParseDate(endDate);

        using var svc = new OutlookCalendarService();
        var result = svc.UpdateEvent(eventId, subject, startDt, endDt, location, body, account);
        return JsonSerializer.Serialize(new { success = result }, JsonOptions);
    }

    [McpServerTool(Name = "delete_event"), Description("Delete a calendar event by its ID.")]
    public string DeleteEvent(
        [Description("Event ID (EntryID from list_events)")] string eventId,
        [Description("Account displayName (from list_accounts). Usually not needed when an event ID is provided.")] string? account = null)
    {
        using var svc = new OutlookCalendarService();
        var result = svc.DeleteEvent(eventId, account);
        return JsonSerializer.Serialize(new { success = result }, JsonOptions);
    }

    [McpServerTool(Name = "find_free_slots"), Description("Find available time slots in the calendar for scheduling. Checks all accounts by default.")]
    public string FindFreeSlots(
        [Description("Start date (yyyy-MM-dd)")] string startDate,
        [Description("End date (yyyy-MM-dd, optional, defaults to 7 days from start)")] string? endDate = null,
        [Description("Slot duration in minutes (optional, defaults to 30)")] int duration = 30,
        [Description("Work day start hour 0-23 (optional, defaults to 9)")] int workDayStart = 9,
        [Description("Work day end hour 0-23 (optional, defaults to 17)")] int workDayEnd = 17,
        [Description("Account displayName to query (from list_accounts, e.g. 'tommy.kihlstrom@thon.no'). Omit to query all accounts.")] string? account = null)
    {
        var start = ParseDate(startDate);
        var end = !string.IsNullOrEmpty(endDate) ? ParseDate(endDate) : start.AddDays(7);

        if (end < start) return Error("End date cannot be before start date.");
        if (workDayEnd <= workDayStart) return Error("Work day end must be after work day start.");

        using var svc = new OutlookCalendarService();
        var slots = svc.FindFreeSlots(start, end, duration, workDayStart, workDayEnd, account);
        return JsonSerializer.Serialize(slots, JsonOptions);
    }

    [McpServerTool(Name = "get_attendee_status"), Description("Check the response status of meeting attendees.")]
    public string GetAttendeeStatus(
        [Description("Event ID (EntryID from list_events)")] string eventId,
        [Description("Account displayName (from list_accounts). Usually not needed when an event ID is provided.")] string? account = null)
    {
        using var svc = new OutlookCalendarService();
        var status = svc.GetAttendeeStatus(eventId, account);
        return JsonSerializer.Serialize(status, JsonOptions);
    }

    [McpServerTool(Name = "get_calendars"), Description("List available calendars in Outlook.")]
    public string GetCalendars()
    {
        using var svc = new OutlookCalendarService();
        var calendars = svc.GetCalendars();
        return JsonSerializer.Serialize(calendars, JsonOptions);
    }

    private static DateTime ParseDate(string dateStr) =>
        DateTime.ParseExact(dateStr, "yyyy-MM-dd", CultureInfo.InvariantCulture);

    private static DateTime ParseDateTime(string dateStr, string timeStr)
    {
        var combined = $"{dateStr} {timeStr}";
        return DateTime.ParseExact(combined, "yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);
    }

    private static string Error(string message) =>
        JsonSerializer.Serialize(new { error = message });
}
