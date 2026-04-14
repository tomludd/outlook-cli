using System.CommandLine;
using System.CommandLine.Invocation;
using System.Globalization;
using System.Text.Json;
using Outlook.COM;

namespace Outlook.Cli;

public static class CalendarCommand
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    public static Command Build()
    {
        var cmd = new Command("calendar", "Manage Outlook calendar events");
        cmd.AddCommand(BuildList());
        cmd.AddCommand(BuildGet());
        cmd.AddCommand(BuildCreate());
        cmd.AddCommand(BuildUpdate());
        cmd.AddCommand(BuildDelete());
        cmd.AddCommand(BuildFreeSlots());
        cmd.AddCommand(BuildAttendees());
        cmd.AddCommand(BuildCalendars());
        return cmd;
    }

    private static Command BuildList()
    {
        var fromArg    = new Argument<string>("from", "Start date yyyy-MM-dd");
        var toArg      = new Argument<string>("to",   "End date yyyy-MM-dd");
        var accountOpt = new Option<string?>("--account", "Account display name (omit for all)");

        var cmd = new Command("list", "List calendar events in a date range") { fromArg, toArg, accountOpt };
        cmd.SetHandler((string from, string to, string? account) =>
        {
            var start = ParseDate(from);
            var end   = ParseDate(to);
            using var svc = new OutlookCalendarService();
            var events = svc.ListEvents(start, end, account);
            Console.WriteLine(JsonSerializer.Serialize(events, JsonOptions));
        }, fromArg, toArg, accountOpt);
        return cmd;
    }

    private static Command BuildGet()
    {
        var idArg = new Argument<string>("id", "Event ID (EntryID from list)");
        var cmd   = new Command("get", "Get full event details including body and attendees") { idArg };
        cmd.SetHandler((string id) =>
        {
            using var svc = new OutlookCalendarService();
            var ev = svc.GetEvent(id);
            Console.WriteLine(JsonSerializer.Serialize(ev, JsonOptions));
        }, idArg);
        return cmd;
    }

    private static Command BuildCreate()
    {
        var subjectOpt    = new Option<string>("--subject",    "Event title") { IsRequired = true };
        var startDateOpt  = new Option<string>("--start-date", "Start date yyyy-MM-dd") { IsRequired = true };
        var startTimeOpt  = new Option<string>("--start-time", "Start time HH:mm") { IsRequired = true };
        var endDateOpt    = new Option<string?>("--end-date",  "End date yyyy-MM-dd (default: same as start)");
        var endTimeOpt    = new Option<string?>("--end-time",  "End time HH:mm (default: 30 min after start)");
        var locationOpt   = new Option<string?>("--location",  "Event location");
        var bodyOpt       = new Option<string?>("--body",      "Event description");
        var meetingOpt    = new Option<bool>("--meeting",      getDefaultValue: () => false, description: "Create as meeting with attendees");
        var attendeesOpt  = new Option<string?>("--attendees", "Attendee emails, semicolon-separated");
        var accountOpt    = new Option<string?>("--account",   "Account display name");

        var cmd = new Command("create", "Create a calendar event") { subjectOpt, startDateOpt, startTimeOpt, endDateOpt, endTimeOpt, locationOpt, bodyOpt, meetingOpt, attendeesOpt, accountOpt };
        cmd.SetHandler((InvocationContext ctx) =>
        {
            var subject   = ctx.ParseResult.GetValueForOption(subjectOpt)!;
            var startDate = ctx.ParseResult.GetValueForOption(startDateOpt)!;
            var startTime = ctx.ParseResult.GetValueForOption(startTimeOpt)!;
            var endDate   = ctx.ParseResult.GetValueForOption(endDateOpt);
            var endTime   = ctx.ParseResult.GetValueForOption(endTimeOpt);
            var location  = ctx.ParseResult.GetValueForOption(locationOpt);
            var body      = ctx.ParseResult.GetValueForOption(bodyOpt);
            var meeting   = ctx.ParseResult.GetValueForOption(meetingOpt);
            var attendees = ctx.ParseResult.GetValueForOption(attendeesOpt);
            var account   = ctx.ParseResult.GetValueForOption(accountOpt);
            var startDt = ParseDateTime(startDate, startTime);
            var endDt = endDate != null && endTime != null ? ParseDateTime(endDate, endTime)
                : endTime != null                          ? ParseDateTime(startDate, endTime)
                : startDt.AddMinutes(30);
            using var svc = new OutlookCalendarService();
            var id = svc.CreateEvent(subject, startDt, endDt, location, body, meeting, attendees, account);
            Console.WriteLine(JsonSerializer.Serialize(new { success = true, eventId = id }, JsonOptions));
        });
        return cmd;
    }

    private static Command BuildUpdate()
    {
        var idArg         = new Argument<string>("id", "Event ID to update");
        var subjectOpt    = new Option<string?>("--subject",    "New title");
        var startDateOpt  = new Option<string?>("--start-date", "New start date yyyy-MM-dd");
        var startTimeOpt  = new Option<string?>("--start-time", "New start time HH:mm");
        var endDateOpt    = new Option<string?>("--end-date",   "New end date yyyy-MM-dd");
        var endTimeOpt    = new Option<string?>("--end-time",   "New end time HH:mm");
        var locationOpt   = new Option<string?>("--location",   "New location");
        var bodyOpt       = new Option<string?>("--body",       "New description");
        var accountOpt    = new Option<string?>("--account",    "Account display name");

        var cmd = new Command("update", "Update an existing calendar event") { idArg, subjectOpt, startDateOpt, startTimeOpt, endDateOpt, endTimeOpt, locationOpt, bodyOpt, accountOpt };
        cmd.SetHandler((InvocationContext ctx) =>
        {
            var id         = ctx.ParseResult.GetValueForArgument(idArg);
            var subject    = ctx.ParseResult.GetValueForOption(subjectOpt);
            var startDate  = ctx.ParseResult.GetValueForOption(startDateOpt);
            var startTime  = ctx.ParseResult.GetValueForOption(startTimeOpt);
            var endDate    = ctx.ParseResult.GetValueForOption(endDateOpt);
            var endTime    = ctx.ParseResult.GetValueForOption(endTimeOpt);
            var location   = ctx.ParseResult.GetValueForOption(locationOpt);
            var body       = ctx.ParseResult.GetValueForOption(bodyOpt);
            var account    = ctx.ParseResult.GetValueForOption(accountOpt);
            DateTime? startDt = startDate != null && startTime != null ? ParseDateTime(startDate, startTime)
                : startDate != null                                     ? ParseDate(startDate)
                : null;
            DateTime? endDt = endDate != null && endTime != null ? ParseDateTime(endDate, endTime)
                : endDate != null                                 ? ParseDate(endDate)
                : null;
            using var svc = new OutlookCalendarService();
            var result = svc.UpdateEvent(id, subject, startDt, endDt, location, body, account);
            Console.WriteLine(JsonSerializer.Serialize(new { success = result }, JsonOptions));
        });
        return cmd;
    }

    private static Command BuildDelete()
    {
        var idArg      = new Argument<string>("id", "Event ID to delete");
        var accountOpt = new Option<string?>("--account", "Account display name");

        var cmd = new Command("delete", "Delete a calendar event") { idArg, accountOpt };
        cmd.SetHandler((string id, string? account) =>
        {
            using var svc = new OutlookCalendarService();
            var result = svc.DeleteEvent(id, account);
            Console.WriteLine(JsonSerializer.Serialize(new { success = result }, JsonOptions));
        }, idArg, accountOpt);
        return cmd;
    }

    private static Command BuildFreeSlots()
    {
        var fromArg        = new Argument<string>("from", "Start date yyyy-MM-dd");
        var toOpt          = new Option<string?>("--to",           "End date yyyy-MM-dd (default: 7 days from start)");
        var durationOpt    = new Option<int>("--duration",         getDefaultValue: () => 30, description: "Slot duration in minutes");
        var workStartOpt   = new Option<int>("--work-start",       getDefaultValue: () => 9,  description: "Work day start hour (0-23)");
        var workEndOpt     = new Option<int>("--work-end",         getDefaultValue: () => 17, description: "Work day end hour (0-23)");
        var accountOpt     = new Option<string?>("--account",      "Account display name (omit for all)");

        var cmd = new Command("free-slots", "Find available time slots for scheduling") { fromArg, toOpt, durationOpt, workStartOpt, workEndOpt, accountOpt };
        cmd.SetHandler((string from, string? to, int duration, int workStart, int workEnd, string? account) =>
        {
            var start = ParseDate(from);
            var end   = to != null ? ParseDate(to) : start.AddDays(7);
            using var svc = new OutlookCalendarService();
            var slots = svc.FindFreeSlots(start, end, duration, workStart, workEnd, account);
            Console.WriteLine(JsonSerializer.Serialize(slots, JsonOptions));
        }, fromArg, toOpt, durationOpt, workStartOpt, workEndOpt, accountOpt);
        return cmd;
    }

    private static Command BuildAttendees()
    {
        var idArg      = new Argument<string>("id", "Event ID");
        var accountOpt = new Option<string?>("--account", "Account display name");

        var cmd = new Command("attendees", "Get attendee response status for a meeting") { idArg, accountOpt };
        cmd.SetHandler((string id, string? account) =>
        {
            using var svc = new OutlookCalendarService();
            var status = svc.GetAttendeeStatus(id, account);
            Console.WriteLine(JsonSerializer.Serialize(status, JsonOptions));
        }, idArg, accountOpt);
        return cmd;
    }

    private static Command BuildCalendars()
    {
        var cmd = new Command("calendars", "List available calendars in Outlook");
        cmd.SetHandler(() =>
        {
            using var svc = new OutlookCalendarService();
            var calendars = svc.GetCalendars();
            Console.WriteLine(JsonSerializer.Serialize(calendars, JsonOptions));
        });
        return cmd;
    }

    private static DateTime ParseDate(string s) =>
        DateTime.ParseExact(s, "yyyy-MM-dd", CultureInfo.InvariantCulture);

    private static DateTime ParseDateTime(string date, string time) =>
        DateTime.ParseExact($"{date} {time}", "yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);
}
