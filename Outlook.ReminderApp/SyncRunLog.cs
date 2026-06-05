namespace Outlook.ReminderApp;

internal sealed class SyncRunLog
{
    public DateTime Timestamp { get; init; }
    public string Message { get; init; } = string.Empty;
    public bool IsError { get; init; }
}
