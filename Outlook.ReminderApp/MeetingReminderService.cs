using System.Diagnostics;
using Outlook.COM;

namespace Outlook.ReminderApp;

internal sealed class MeetingReminderService : IDisposable
{
    private static readonly TimeSpan UpcomingWindow = TimeSpan.FromMinutes(10);
    private static readonly TimeSpan AutoOpenGraceWindow = TimeSpan.FromMinutes(2);
    private static readonly TimeSpan QueryHistoryWindow = TimeSpan.FromHours(8);
    private static readonly TimeSpan QueryFutureWindow = TimeSpan.FromHours(8);

    private readonly MeetingActionStateStore _stateStore = new();

    public IReadOnlyList<ReminderMeeting> GetVisibleMeetings(DateTime now)
    {
        _stateStore.Cleanup(now);
        var visibleWindowEnd = now.Add(UpcomingWindow);
        var allMeetings = GetMeetings(now);

        var visible = allMeetings
            .Where(x => x.IsOngoing(now) || (x.Start >= now && x.Start <= visibleWindowEnd))
            .Where(x => !_stateStore.IsDismissed(x.Id, now))
            .ToList();

        foreach (var meeting in visible)
        {
            meeting.IsOverlapping = visible.Any(other =>
                !string.Equals(other.Id, meeting.Id, StringComparison.OrdinalIgnoreCase) &&
                IsOverlapping(meeting, other));
        }

        visible = visible
            .OrderBy(x => x.IsOngoing(now) ? 0 : 1)
            .ThenBy(x => x.Start)
            .ToList();

        return visible;
    }

    public ReminderMeeting? GetNextMeeting(DateTime now)
    {
        _stateStore.Cleanup(now);

        return GetMeetings(now)
            .Where(x => x.End > now)
            .OrderBy(x => x.IsOngoing(now) ? 0 : 1)
            .ThenBy(x => x.Start)
            .FirstOrDefault();
    }

    public bool IsDismissed(string meetingId, DateTime now)
    {
        return _stateStore.IsDismissed(meetingId, now);
    }

    public void Dismiss(ReminderMeeting meeting)
    {
        _stateStore.MarkDismissed(meeting.Id, meeting.End);
    }

    public void TryAutoOpenDueMeetings(IEnumerable<ReminderMeeting> meetings, DateTime now)
    {
        foreach (var meeting in meetings)
        {
            if (!meeting.HasTeamsJoinUrl)
            {
                continue;
            }

            if (meeting.IsDeclined)
            {
                continue;
            }

            if (_stateStore.IsDismissed(meeting.Id, now))
            {
                continue;
            }

            if (_stateStore.IsAutoOpened(meeting.Id, now))
            {
                continue;
            }

            if (now < meeting.Start)
            {
                continue;
            }

            if (now > meeting.Start.Add(AutoOpenGraceWindow))
            {
                continue;
            }

            OpenMeetingUrl(meeting.TeamsJoinUrl!);
            _stateStore.MarkAutoOpened(meeting.Id, meeting.End);
        }
    }

    public void OpenJoin(ReminderMeeting meeting)
    {
        if (!meeting.HasTeamsJoinUrl)
        {
            return;
        }

        OpenMeetingUrl(meeting.TeamsJoinUrl!);
    }

    private static ReminderMeeting? ToReminderMeeting(Dictionary<string, object?> row)
    {
        var id = GetValue(row, "id");
        var subject = GetValue(row, "subject");
        var startText = GetValue(row, "start");
        var endText = GetValue(row, "end");

        if (string.IsNullOrWhiteSpace(id) || string.IsNullOrWhiteSpace(subject) ||
            !DateTime.TryParse(startText, out var start) || !DateTime.TryParse(endText, out var end))
        {
            return null;
        }

        var location = GetValue(row, "location") ?? string.Empty;
        var body = GetValue(row, "body") ?? string.Empty;
        var responseStatus = GetValue(row, "responseStatus") ?? "Unknown";

        bool isMeeting = false;
        if (row.TryGetValue("isMeeting", out var isMeetingRaw) && isMeetingRaw is bool b)
        {
            isMeeting = b;
        }

        bool isCancelled = false;
        if (row.TryGetValue("isCancelled", out var isCancelledRaw) && isCancelledRaw is bool c)
        {
            isCancelled = c;
        }

        bool isResponseRequested = false;
        if (row.TryGetValue("responseRequested", out var responseRequestedRaw) && responseRequestedRaw is bool responseRequested)
        {
            isResponseRequested = responseRequested;
        }

        var teamsJoinUrl = TeamsJoinLinkResolver.Resolve(body, location);

        return new ReminderMeeting
        {
            Id = id,
            Subject = subject,
            Start = start,
            End = end,
            Location = location,
            Body = body,
            IsMeeting = isMeeting,
            IsCancelled = isCancelled,
            IsResponseRequested = isResponseRequested,
            ResponseStatus = responseStatus,
            TeamsJoinUrl = teamsJoinUrl
        };
    }

    private static List<ReminderMeeting> GetMeetings(DateTime now)
    {
        using var calendarService = new OutlookCalendarService();
        var events = calendarService.ListEvents(now.Subtract(QueryHistoryWindow), now.Add(QueryFutureWindow), account: null);

        return events
            .Select(ToReminderMeeting)
            .Where(x => x is not null && !x.Body.Contains("[outlook-sync:") && !x.IsCancelled)
            .Cast<ReminderMeeting>()
            .DistinctBy(x => x.Id)
            .ToList();
    }

    private static string? GetValue(Dictionary<string, object?> row, string key)
    {
        return row.TryGetValue(key, out var value) ? value?.ToString() : null;
    }

    private static void OpenMeetingUrl(string url)
    {
        Process.Start(new ProcessStartInfo
        {
            FileName = url,
            UseShellExecute = true
        });
    }

    private static bool IsOverlapping(ReminderMeeting left, ReminderMeeting right)
    {
        return left.Start < right.End && left.End > right.Start;
    }

    public void Dispose()
    {
    }
}