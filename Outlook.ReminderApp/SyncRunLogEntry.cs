namespace Outlook.ReminderApp;

internal sealed class SyncRunLogEntry
{
    public DateTime Timestamp { get; init; }
    public string RuleLabel { get; init; } = string.Empty;
    public string Status { get; init; } = string.Empty;
    public string Message { get; init; } = string.Empty;

    public override string ToString()
    {
        var rulePart = string.IsNullOrWhiteSpace(RuleLabel) ? "" : $" [{RuleLabel}]";
        var messagePart = string.IsNullOrWhiteSpace(Message) ? "" : $" - {Message}";
        return $"{Timestamp:HH:mm:ss} {Status}{rulePart}{messagePart}";
    }
}
