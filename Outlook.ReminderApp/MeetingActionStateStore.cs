namespace Outlook.ReminderApp;

internal sealed class MeetingActionStateStore
{
    private readonly Dictionary<string, DateTime> _dismissedUntil = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, DateTime> _openedUntil = new(StringComparer.OrdinalIgnoreCase);

    public void MarkDismissed(string meetingId, DateTime end)
    {
        _dismissedUntil[meetingId] = end;
    }

    public bool IsDismissed(string meetingId, DateTime now)
    {
        return _dismissedUntil.TryGetValue(meetingId, out var until) && until > now;
    }

    public bool IsAutoOpened(string meetingId, DateTime now)
    {
        return _openedUntil.TryGetValue(meetingId, out var until) && until > now;
    }

    public void MarkAutoOpened(string meetingId, DateTime end)
    {
        _openedUntil[meetingId] = end;
    }

    public void Cleanup(DateTime now)
    {
        CleanupMap(_dismissedUntil, now);
        CleanupMap(_openedUntil, now);
    }

    private static void CleanupMap(Dictionary<string, DateTime> values, DateTime now)
    {
        if (values.Count == 0)
        {
            return;
        }

        var expiredKeys = values.Where(x => x.Value <= now).Select(x => x.Key).ToArray();
        foreach (var key in expiredKeys)
        {
            values.Remove(key);
        }
    }
}