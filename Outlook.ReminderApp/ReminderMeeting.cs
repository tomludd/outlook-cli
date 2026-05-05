namespace Outlook.ReminderApp;

internal sealed class ReminderMeeting
{
    public required string Id { get; init; }
    public required string Subject { get; init; }
    public required DateTime Start { get; init; }
    public required DateTime End { get; init; }
    public string Location { get; init; } = string.Empty;
    public string Body { get; init; } = string.Empty;
    public bool IsMeeting { get; init; }
    public bool IsCancelled { get; init; }
    public bool IsResponseRequested { get; init; }
    public string ResponseStatus { get; init; } = "Unknown";
    public string? TeamsJoinUrl { get; init; }
    public bool IsOverlapping { get; set; }

    public bool IsOngoing(DateTime now)
    {
        return Start <= now && End > now;
    }

    public bool IsDeclined => string.Equals(ResponseStatus, "Declined", StringComparison.OrdinalIgnoreCase);
    public bool IsNotResponded =>
        IsMeeting &&
        IsResponseRequested &&
        (string.Equals(ResponseStatus, "Not Responded", StringComparison.OrdinalIgnoreCase) ||
         string.Equals(ResponseStatus, "NotResponded", StringComparison.OrdinalIgnoreCase));

    public bool HasTeamsJoinUrl => !string.IsNullOrWhiteSpace(TeamsJoinUrl);

    public string DisplaySubject =>
        Subject.StartsWith("Following: ", StringComparison.Ordinal)
            ? "Follow: " + Subject[11..]
            : Subject;
}