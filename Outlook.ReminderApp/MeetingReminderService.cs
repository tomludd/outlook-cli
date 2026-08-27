using System.Diagnostics;
using System.Runtime.InteropServices;
using Outlook.COM;

namespace Outlook.ReminderApp;

internal sealed class MeetingReminderService
{
    private static readonly TimeSpan UpcomingWindow = TimeSpan.FromMinutes(10);
    private static readonly TimeSpan AutoOpenGraceWindow = TimeSpan.FromMinutes(2);

    private readonly MeetingActionStateStore _stateStore = new();
    private readonly OutlookCalendarService _calendarService = new();

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
            .Where(x => !x.IsOutlookSynced)
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

    /// <summary>
    /// Returns 0 if the meeting's account domain matches the organizer's email domain (best match),
    /// 1 otherwise. Used to prefer the copy of a duplicated meeting from the organizer's own org.
    /// </summary>
    private static int OrganizerDomainMatchScore(ReminderMeeting m)
    {
        if (string.IsNullOrEmpty(m.OrganizerEmail) || string.IsNullOrEmpty(m.Account))
            return 1;
        var orgDomain = DomainOf(m.OrganizerEmail);
        var accDomain = DomainOf(m.Account);
        return string.Equals(orgDomain, accDomain, StringComparison.OrdinalIgnoreCase) ? 0 : 1;
    }

    private static string DomainOf(string email)
    {
        var at = email.IndexOf('@');
        return at >= 0 ? email[(at + 1)..] : email;
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
        var organizerEmail = GetValue(row, "organizerEmail") ?? string.Empty;

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

        bool isAllDay = false;
        if (row.TryGetValue("isAllDay", out var isAllDayRaw) && isAllDayRaw is bool ad)
        {
            isAllDay = ad;
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
            IsAllDay = isAllDay,
            IsResponseRequested = isResponseRequested,
            ResponseStatus = responseStatus,
            TeamsJoinUrl = teamsJoinUrl,
            TeamsChatUrl = teamsChatUrl,
            Account = account,
            OrganizerEmail = organizerEmail
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
            .Where(x => x.Start < todayEnd && x.End > todayStart && !x.IsOutlookSynced)
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
        var events = _calendarService.ListEvents(from, to, account: null);

        return events
            .Select(ToReminderMeeting)
            .Where(x => x is not null)
            .Cast<ReminderMeeting>()
            .DistinctBy(x => x.Id)
            .GroupBy(x => (x.Subject, x.Start, x.End))
            .Select(g => g
                .OrderBy(x => OrganizerDomainMatchScore(x))
                .ThenBy(x => x.IsDeclined ? 2 : x.IsNotResponded ? 1 : 0)
                .ThenByDescending(x => x.HasTeamsJoinUrl)
                .First())
            .ToList();
    }

    /// <summary>
    /// Responds to a meeting asynchronously so the UI thread is never blocked by a slow Outlook.
    /// The COM call runs on the persistent STA worker via ComTimeout.Run; Task.Run ensures the
    /// blocking wait happens off the UI thread.
    /// </summary>
    public Task RespondToMeetingAsync(string meetingId, bool accept)
        => Task.Run(() => _calendarService.RespondToMeeting(meetingId, accept ? 3 : 4));

    /// <summary>
    /// Fetches meeting details asynchronously so the UI thread is never blocked by a slow Outlook.
    /// </summary>
    public Task<MeetingDetails?> GetMeetingDetailsAsync(string meetingId)
        => Task.Run(() =>
        {
            try
            {
                var dict = _calendarService.GetEvent(meetingId);

                var subject  = dict.TryGetValue("subject",  out var s)  ? s?.ToString()  ?? string.Empty : string.Empty;
                var organizer = dict.TryGetValue("organizer", out var o)  ? o?.ToString()  ?? string.Empty : string.Empty;
                var location  = dict.TryGetValue("location",  out var l)  ? l?.ToString()  ?? string.Empty : string.Empty;
                var body      = dict.TryGetValue("body",      out var bd) ? bd?.ToString() ?? string.Empty : string.Empty;
                var htmlBody  = dict.TryGetValue("htmlBody",  out var hb) ? hb?.ToString() ?? string.Empty : string.Empty;

                DateTime start = DateTime.MinValue, end = DateTime.MinValue;
                if (dict.TryGetValue("start", out var sv) && sv is string startStr)
                    DateTime.TryParseExact(startStr, "yyyy-MM-dd HH:mm", null, System.Globalization.DateTimeStyles.None, out start);
                if (dict.TryGetValue("end", out var ev) && ev is string endStr)
                    DateTime.TryParseExact(endStr, "yyyy-MM-dd HH:mm", null, System.Globalization.DateTimeStyles.None, out end);

                var attendees = new List<AttendeeInfo>();
                if (dict.TryGetValue("attendees", out var raw) && raw is List<Dictionary<string, string>> list)
                {
                    foreach (var a in list)
                    {
                        var name   = a.TryGetValue("name",           out var n) ? n : string.Empty;
                        var status = a.TryGetValue("responseStatus", out var r) ? r : "Unknown";
                        var email  = a.TryGetValue("email",          out var e) ? e : string.Empty;
                        attendees.Add(new AttendeeInfo(name, status, email));
                    }
                }

                return new MeetingDetails(subject, start, end, organizer, location, body, htmlBody, attendees);
            }
            catch
            {
                return null;
            }
        });

    /// <summary>
    /// Opens the meeting item in Outlook asynchronously so the UI thread is never blocked.
    /// </summary>
    public Task OpenInOutlookAsync(string meetingId)
        => Task.Run(() =>
        {
            try { _calendarService.OpenItem(meetingId); }
            catch { }
        });

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
        _ = TryBringTeamsToFrontAsync();
    }

    private static async Task TryBringTeamsToFrontAsync()
    {
        var deadline = DateTime.UtcNow.AddSeconds(30);
        while (DateTime.UtcNow < deadline)
        {
            await Task.Delay(500);
            var hwnd = FindTeamsWindow();
            if (hwnd != IntPtr.Zero)
            {
                ShowWindow(hwnd, SW_RESTORE);
                SetForegroundWindow(hwnd);
                return;
            }
        }
    }

    private static IntPtr FindTeamsWindow()
    {
        foreach (var name in (string[])["ms-teams", "Teams"])
        {
            foreach (var p in Process.GetProcessesByName(name))
            {
                using (p)
                {
                    if (p.MainWindowHandle != IntPtr.Zero)
                        return p.MainWindowHandle;
                }
            }
        }
        return IntPtr.Zero;
    }

    private const int SW_RESTORE = 9;

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    private static bool IsOverlapping(ReminderMeeting left, ReminderMeeting right)
    {
        return left.Start < right.End && left.End > right.Start;
    }

}

internal sealed record AttendeeInfo(string Name, string ResponseStatus, string Email = "");

internal sealed record MeetingDetails(
    string Subject,
    DateTime Start,
    DateTime End,
    string Organizer,
    string Location,
    string Body,
    string HtmlBody,
    IReadOnlyList<AttendeeInfo> Attendees);