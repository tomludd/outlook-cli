using Outlook.COM;

namespace Outlook.Cli;

public enum SyncMode
{
    /// <summary>Creates anonymous "Busy" blocking events in the target calendar.</summary>
    Block,
    /// <summary>Copies title and description of events to the target calendar.</summary>
    Copy,
}

public class CalendarSyncService
{
    // Markers embedded in synced event descriptions so we can identify and manage them.
    // Format: [outlook-sync:block:<hash>] or [outlook-sync:copy:<hash>]

    // Outlook busy status constants
    private const int OlBusy = 2;
    private const int OlFree = 0;
    private const int OlOutOfOffice = 3;

    public SyncSummary RunSync(
        string sourceAccount,
        string targetAccount,
        DateTime from,
        DateTime to,
        SyncMode mode = SyncMode.Block,
        bool outsideWorkHoursOnly = false,
        int workDayStartHour = 7,
        int workDayEndHour = 18)
    {
        var summary = new SyncSummary();
        var modeKey = mode == SyncMode.Copy ? "copy" : "block";
        var marker = $"[outlook-sync:{modeKey}:{ComputeRuleId(sourceAccount, targetAccount)}]";

        using var calService = new OutlookCalendarService();

        var sourceEvents = calService.ListEvents(from, to, sourceAccount)
            .Where(ShouldSync)
            .Where(e => !outsideWorkHoursOnly || IsOutsideWorkHours(e, workDayStartHour, workDayEndHour))
            .ToList();

        var allTargetEvents = calService.ListEvents(from, to, targetAccount);

        var ourSyncedEvents = allTargetEvents
            .Where(e => HasMarker(e, marker))
            .ToList();

        var realTargetSlots = allTargetEvents
            .Where(e => !HasMarker(e, marker))
            .Select(e => (Start: ParseDate(e["start"]), End: ParseDate(e["end"])))
            .ToList();

        var sourceSlots = sourceEvents
            .Select(e => (Start: ParseDate(e["start"]), End: ParseDate(e["end"])))
            .ToHashSet();

        // Delete synced events whose source slot no longer exists in the window
        var syncedSlots = new HashSet<(DateTime Start, DateTime End)>();
        foreach (var synced in ourSyncedEvents)
        {
            var slot = (Start: ParseDate(synced["start"]), End: ParseDate(synced["end"]));
            if (!sourceSlots.Contains(slot))
            {
                try { calService.DeleteEvent((string)synced["id"]!, targetAccount); }
                catch { /* already gone from target */ }
                summary.Deleted++;
            }
            else
            {
                syncedSlots.Add(slot);
            }
        }

        // Create synced events for source slots not yet present in target
        foreach (var srcEvent in sourceEvents)
        {
            var srcStart = ParseDate(srcEvent["start"]);
            var srcEnd = ParseDate(srcEvent["end"]);
            var slot = (Start: srcStart, End: srcEnd);

            if (syncedSlots.Contains(slot)) { summary.Skipped++; continue; }
            if (mode == SyncMode.Block && TargetAlreadyHasEvent(realTargetSlots, srcStart, srcEnd)) { summary.Skipped++; continue; }

            var srcBusyStatus = (string?)srcEvent["busyStatus"];

            string subject;
            string body;
            int busyStatus;

            if (mode == SyncMode.Copy)
            {
                subject = (string?)srcEvent["subject"] ?? string.Empty;
                var srcBody = srcEvent.TryGetValue("body", out var b) ? b as string ?? string.Empty : string.Empty;
                body = string.IsNullOrEmpty(srcBody) ? marker : $"{srcBody}\n{marker}";
                busyStatus = srcBusyStatus == "Out of Office" ? OlOutOfOffice : OlFree;
            }
            else
            {
                subject = srcBusyStatus == "Out of Office" ? "Out of Office" : "Busy";
                body = $"Blocked from {sourceAccount}\n{marker}";
                busyStatus = srcBusyStatus == "Out of Office" ? OlOutOfOffice : OlBusy;
            }

            try
            {
                calService.CreateEvent(
                    subject: subject,
                    startDateTime: srcStart,
                    endDateTime: srcEnd,
                    location: null,
                    body: body,
                    isMeeting: false,
                    attendees: null,
                    account: targetAccount,
                    reminderEnabled: false,
                    busyStatus: busyStatus);
                summary.Created++;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"  Failed to create event for {srcStart:yyyy-MM-dd HH:mm}: {ex.Message}");
                summary.Skipped++;
            }
        }

        return summary;
    }

    private static bool IsOutsideWorkHours(Dictionary<string, object?> ev, int startHour, int endHour)
    {
        var start = ParseDate(ev["start"]);
        var end = ParseDate(ev["end"]);
        return start.TimeOfDay < TimeSpan.FromHours(startHour)
            || end.TimeOfDay > TimeSpan.FromHours(endHour);
    }

    // Derives a stable 8-char hex ID from the source:target pair so markers are consistent across runs.
    private static string ComputeRuleId(string source, string target)
    {
        var bytes = System.Security.Cryptography.SHA256.HashData(
            System.Text.Encoding.UTF8.GetBytes($"{source}:{target}"));
        return Convert.ToHexString(bytes)[..8].ToLowerInvariant();
    }

    private static bool HasMarker(Dictionary<string, object?> ev, string marker)
    {
        return ev.TryGetValue("body", out var body) && body is string bodyStr && bodyStr.Contains(marker);
    }

    private static bool ShouldSync(Dictionary<string, object?> ev)
    {
        var busyStatus = (string?)ev["busyStatus"];
        if (busyStatus is not ("Busy" or "Out of Office")) return false;

        // Skip blocking events created by this tool — avoids cascading syncs
        if (ev.TryGetValue("body", out var body) && body is string bodyStr && bodyStr.Contains("[outlook-sync:"))
            return false;

        return true;
    }

    private static bool TargetAlreadyHasEvent(
        List<(DateTime Start, DateTime End)> targetEvents,
        DateTime srcStart, DateTime srcEnd)
    {
        return targetEvents.Any(t => t.Start == srcStart && t.End == srcEnd);
    }

    private static DateTime ParseDate(object? value)
    {
        if (value is DateTime dt) return dt;
        if (value is string s && DateTime.TryParseExact(s, "yyyy-MM-dd HH:mm", null,
                System.Globalization.DateTimeStyles.None, out var parsed))
            return parsed;
        return DateTime.MinValue;
    }
}

public class SyncSummary
{
    public int Created { get; set; }
    public int Deleted { get; set; }
    public int Skipped { get; set; }
    public int Errors { get; set; }
}
