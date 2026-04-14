using System.CommandLine;
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
        cmd.Subcommands.Add(BuildList());
        cmd.Subcommands.Add(BuildGet());
        cmd.Subcommands.Add(BuildCreate());
        cmd.Subcommands.Add(BuildUpdate());
        cmd.Subcommands.Add(BuildDelete());
        cmd.Subcommands.Add(BuildFreeSlots());
        cmd.Subcommands.Add(BuildAttendees());
        cmd.Subcommands.Add(BuildCalendars());
        return cmd;
    }

    private static Command BuildList()
    {
        var fromArg    = new Argument<string>("from") { Description = "Start date yyyy-MM-dd" };
        var toArg      = new Argument<string>("to")   { Description = "End date yyyy-MM-dd" };
        var accountOpt = new Option<string?>("--account") { Description = "Account display name (omit for all)" };

        var cmd = new Command("list", "List calendar events in a date range");
        cmd.Arguments.Add(fromArg); cmd.Arguments.Add(toArg); cmd.Options.Add(accountOpt);
        cmd.SetAction(ctx =>
        {
            var start   = ParseDate(ctx.GetValue(fromArg)!);
            var end     = ParseDate(ctx.GetValue(toArg)!);
            var account = ctx.GetValue(accountOpt);
            using var svc = new OutlookCalendarService();
            var events = svc.ListEvents(start, end, account);
            Console.WriteLine(JsonSerializer.Serialize(events, JsonOptions));
        });
        return cmd;
    }

    private static Command BuildGet()
    {
        var idArg = new Argument<string>("id") { Description = "Event ID (EntryID from list)" };
        var cmd   = new Command("get", "Get full event details including body and attendees");
        cmd.Arguments.Add(idArg);
        cmd.SetAction(ctx =>
        {
            var id = ctx.GetValue(idArg)!;
            using var svc = new OutlookCalendarService();
            var ev = svc.GetEvent(id);
            Console.WriteLine(JsonSerializer.Serialize(ev, JsonOptions));
        });
        return cmd;
    }

    private static Command BuildCreate()
    {
        var subjectOpt   = new Option<string>("--subject") { Description = "Event title", Required = true };
        var startDateOpt = new Option<string>("--start-date") { Description = "Start date yyyy-MM-dd", Required = true };
        var startTimeOpt = new Option<string>("--start-time") { Description = "Start time HH:mm", Required = true };
        var endDateOpt   = new Option<string?>("--end-date") { Description = "End date yyyy-MM-dd (default: same as start)" };
        var endTimeOpt   = new Option<string?>("--end-time") { Description = "End time HH:mm (default: 30 min after start)" };
        var locationOpt  = new Option<string?>("--location") { Description = "Event location" };
        var bodyOpt      = new Option<string?>("--body") { Description = "Event description" };
        var meetingOpt   = new Option<bool>("--meeting") { Description = "Create as meeting with attendees", DefaultValueFactory = _ => false };
        var attendeesOpt = new Option<string?>("--attendees") { Description = "Attendee emails, semicolon-separated" };
        var accountOpt   = new Option<string?>("--account") { Description = "Account display name" };

        var cmd = new Command("create", "Create a calendar event");
        cmd.Options.Add(subjectOpt); cmd.Options.Add(startDateOpt); cmd.Options.Add(startTimeOpt);
        cmd.Options.Add(endDateOpt); cmd.Options.Add(endTimeOpt); cmd.Options.Add(locationOpt);
        cmd.Options.Add(bodyOpt); cmd.Options.Add(meetingOpt); cmd.Options.Add(attendeesOpt); cmd.Options.Add(accountOpt);
        cmd.SetAction(ctx =>
        {
            var subject   = ctx.GetValue(subjectOpt)!;
            var startDate = ctx.GetValue(startDateOpt)!;
            var startTime = ctx.GetValue(startTimeOpt)!;
            var endDate   = ctx.GetValue(endDateOpt);
            var endTime   = ctx.GetValue(endTimeOpt);
            var location  = ctx.GetValue(locationOpt);
            var body      = ctx.GetValue(bodyOpt);
            var meeting   = ctx.GetValue(meetingOpt);
            var attendees = ctx.GetValue(attendeesOpt);
            var account   = ctx.GetValue(accountOpt);
            var startDt = ParseDateTime(startDate, startTime);
            var endDt   = endDate != null && endTime != null ? ParseDateTime(endDate, endTime)
                        : endTime != null                    ? ParseDateTime(startDate, endTime)
                        : startDt.AddMinutes(30);
            using var svc = new OutlookCalendarService();
            var id = svc.CreateEvent(subject, startDt, endDt, location, body, meeting, attendees, account);
            Console.WriteLine(JsonSerializer.Serialize(new { success = true, eventId = id }, JsonOptions));
        });
        return cmd;
    }

    private static Command BuildUpdate()
    {
        var idArg        = new Argument<string>("id") { Description = "Event ID to update" };
        var subjectOpt   = new Option<string?>("--subject") { Description = "New title" };
        var startDateOpt = new Option<string?>("--start-date") { Description = "New start date yyyy-MM-dd" };
        var startTimeOpt = new Option<string?>("--start-time") { Description = "New start time HH:mm" };
        var endDateOpt   = new Option<string?>("--end-date") { Description = "New end date yyyy-MM-dd" };
        var endTimeOpt   = new Option<string?>("--end-time") { Description = "New end time HH:mm" };
        var locationOpt  = new Option<string?>("--location") { Description = "New location" };
        var bodyOpt      = new Option<string?>("--body") { Description = "New description" };
        var accountOpt   = new Option<string?>("--account") { Description = "Account display name" };

        var cmd = new Command("update", "Update an existing calendar event");
        cmd.Arguments.Add(idArg);
        cmd.Options.Add(subjectOpt); cmd.Options.Add(startDateOpt); cmd.Options.Add(startTimeOpt);
        cmd.Options.Add(endDateOpt); cmd.Options.Add(endTimeOpt); cmd.Options.Add(locationOpt);
        cmd.Options.Add(bodyOpt); cmd.Options.Add(accountOpt);
        cmd.SetAction(ctx =>
        {
            var id        = ctx.GetValue(idArg)!;
            var subject   = ctx.GetValue(subjectOpt);
            var startDate = ctx.GetValue(startDateOpt);
            var startTime = ctx.GetValue(startTimeOpt);
            var endDate   = ctx.GetValue(endDateOpt);
            var endTime   = ctx.GetValue(endTimeOpt);
            var location  = ctx.GetValue(locationOpt);
            var body      = ctx.GetValue(bodyOpt);
            var account   = ctx.GetValue(accountOpt);
            DateTime? startDt = startDate != null && startTime != null ? ParseDateTime(startDate, startTime)
                              : startDate != null                       ? ParseDate(startDate)
                              : null;
            DateTime? endDt = endDate != null && endTime != null ? ParseDateTime(endDate, endTime)
                            : endDate != null                     ? ParseDate(endDate)
                            : null;
            using var svc = new OutlookCalendarService();
            var result = svc.UpdateEvent(id, subject, startDt, endDt, location, body, account);
            Console.WriteLine(JsonSerializer.Serialize(new { success = result }, JsonOptions));
        });
        return cmd;
    }

    private static Command BuildDelete()
    {
        var idArg      = new Argument<string>("id") { Description = "Event ID to delete" };
        var accountOpt = new Option<string?>("--account") { Description = "Account display name" };

        var cmd = new Command("delete", "Delete a calendar event");
        cmd.Arguments.Add(idArg); cmd.Options.Add(accountOpt);
        cmd.SetAction(ctx =>
        {
            var id      = ctx.GetValue(idArg)!;
            var account = ctx.GetValue(accountOpt);
            using var svc = new OutlookCalendarService();
            var result = svc.DeleteEvent(id, account);
            Console.WriteLine(JsonSerializer.Serialize(new { success = result }, JsonOptions));
        });
        return cmd;
    }

    private static Command BuildFreeSlots()
    {
        var fromArg      = new Argument<string>("from") { Description = "Start date yyyy-MM-dd" };
        var toOpt        = new Option<string?>("--to") { Description = "End date yyyy-MM-dd (default: 7 days from start)" };
        var durationOpt  = new Option<int>("--duration") { Description = "Slot duration in minutes", DefaultValueFactory = _ => 30 };
        var workStartOpt = new Option<int>("--work-start") { Description = "Work day start hour (0-23)", DefaultValueFactory = _ => 9 };
        var workEndOpt   = new Option<int>("--work-end") { Description = "Work day end hour (0-23)", DefaultValueFactory = _ => 17 };
        var accountOpt   = new Option<string?>("--account") { Description = "Account display name (omit for all)" };

        var cmd = new Command("free-slots", "Find available time slots for scheduling");
        cmd.Arguments.Add(fromArg);
        cmd.Options.Add(toOpt); cmd.Options.Add(durationOpt); cmd.Options.Add(workStartOpt);
        cmd.Options.Add(workEndOpt); cmd.Options.Add(accountOpt);
        cmd.SetAction(ctx =>
        {
            var start     = ParseDate(ctx.GetValue(fromArg)!);
            var toVal     = ctx.GetValue(toOpt);
            var end       = toVal != null ? ParseDate(toVal) : start.AddDays(7);
            var duration  = ctx.GetValue(durationOpt);
            var workStart = ctx.GetValue(workStartOpt);
            var workEnd   = ctx.GetValue(workEndOpt);
            var account   = ctx.GetValue(accountOpt);
            using var svc = new OutlookCalendarService();
            var slots = svc.FindFreeSlots(start, end, duration, workStart, workEnd, account);
            Console.WriteLine(JsonSerializer.Serialize(slots, JsonOptions));
        });
        return cmd;
    }

    private static Command BuildAttendees()
    {
        var idArg      = new Argument<string>("id") { Description = "Event ID" };
        var accountOpt = new Option<string?>("--account") { Description = "Account display name" };

        var cmd = new Command("attendees", "Get attendee response status for a meeting");
        cmd.Arguments.Add(idArg); cmd.Options.Add(accountOpt);
        cmd.SetAction(ctx =>
        {
            var id      = ctx.GetValue(idArg)!;
            var account = ctx.GetValue(accountOpt);
            using var svc = new OutlookCalendarService();
            var status = svc.GetAttendeeStatus(id, account);
            Console.WriteLine(JsonSerializer.Serialize(status, JsonOptions));
        });
        return cmd;
    }

    private static Command BuildCalendars()
    {
        var cmd = new Command("calendars", "List available calendars in Outlook");
        cmd.SetAction(_ =>
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




