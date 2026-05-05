using System.Diagnostics;
using Outlook.COM;

namespace Outlook.ReminderApp;

internal sealed class MeetingReminderService : IDisposable
{
    private static readonly TimeSpan UpcomingWindow = TimeSpan.FromMinutes(10);
    private static readonly TimeSpan AutoOpenGraceWindow = TimeSpan.FromMinutes(2);

    private readonly MeetingActionStateStore _stateStore = new();

    /// <summary>
    /// Returns meetings visible in the notification widget, derived from the supplied cached list.
    /// Excludes cancelled meetings, dismissed meetings, and meetings outside the upcoming window.
    /// </summary>
    public IReadOnlyList<ReminderMeeting> GetVisibleMeetings(DateTime now, IReadOnlyList<ReminderMeeting> allMeetings)
    {
        _stateStore.Cleanup(now);
        var visibleWindowEnd = now.Add(UpcomingWindow);

        var visible = allMeetings
            .Where(x => !x.IsCancelled)
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

    /// <summary>
    /// Returns the next upcoming or ongoing meeting (non-cancelled), derived from the supplied cached list.
    /// </summary>
    public ReminderMeeting? GetNextMeeting(DateTime now, IReadOnlyList<ReminderMeeting> allMeetings)
    {
        _stateStore.Cleanup(now);

        return allMeetings
            .Where(x => !x.IsCancelled && x.End > now)
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
        var account = GetValue(row, "account") ?? string.Empty;

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
        var teamsChatUrl = TeamsJoinLinkResolver.ResolveChat(body, location)
                           ?? TeamsJoinLinkResolver.DeriveChatUrlFromJoinUrl(teamsJoinUrl)
                           ?? TeamsJoinLinkResolver.DeriveChatUrlFromDecodedBody(body);

        // Append accountHint so Teams opens the meeting/chat with the correct account
        // when multiple accounts are signed in.
        if (!string.IsNullOrEmpty(account))
        {
            var hint = Uri.EscapeDataString(account);
            if (teamsJoinUrl is not null)
            {
                var sep = teamsJoinUrl.Contains('?') ? '&' : '?';
                teamsJoinUrl = $"{teamsJoinUrl}{sep}accountHint={hint}";
            }
            if (teamsChatUrl is not null)
            {
                var sep = teamsChatUrl.Contains('?') ? '&' : '?';
                teamsChatUrl = $"{teamsChatUrl}{sep}accountHint={hint}";
            }
        }

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
            TeamsJoinUrl = teamsJoinUrl,
            TeamsChatUrl = teamsChatUrl,
            Account = account
        };
    }

    /// <summary>
    /// Returns all of today's meetings (including cancelled) derived from the supplied cached list.
    /// </summary>
    public IReadOnlyList<ReminderMeeting> GetTodaysMeetings(DateTime now, IReadOnlyList<ReminderMeeting> allMeetings)
    {
        var todayStart = now.Date;
        var todayEnd   = todayStart.AddDays(1);

        return allMeetings
            .Where(x => x.Start >= todayStart && x.Start < todayEnd && !x.Body.Contains("[outlook-sync:"))
            .DistinctBy(x => x.Id)
            .OrderBy(x => x.Start)
            .ToList();
    }

    /// <summary>
    /// Fetches all meetings (including cancelled) from Outlook COM for the given time range.
    /// Called by <see cref="MeetingCache"/> on the UI/STA thread.
    /// </summary>
    public IReadOnlyList<ReminderMeeting> FetchAll(DateTime from, DateTime to)
    {
        using var calendarService = new OutlookCalendarService();
        var events = calendarService.ListEvents(from, to, account: null);

        return events
            .Select(ToReminderMeeting)
            .Where(x => x is not null)
            .Cast<ReminderMeeting>()
            .DistinctBy(x => x.Id)
            .ToList();
    }

    public void RespondToMeeting(string meetingId, bool accept)
    {
        using var calendarService = new OutlookCalendarService();
        calendarService.RespondToMeeting(meetingId, accept ? 3 : 4);
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