namespace Outlook.ReminderApp;

/// <summary>
/// Maintains a periodically refreshed in-memory snapshot of meetings fetched from Outlook.
/// Must be accessed from the UI (STA) thread — uses a WinForms timer internally so all
/// callbacks fire on the message loop without any cross-thread concerns.
/// </summary>
internal sealed class MeetingCache : IDisposable
{
    private static readonly TimeSpan QueryHistoryWindow = TimeSpan.FromHours(8);
    private static readonly TimeSpan QueryFutureWindow  = TimeSpan.FromHours(8);

    private readonly MeetingReminderService _service;
    private readonly System.Windows.Forms.Timer _timer;

    /// <summary>
    /// All fetched meetings (including cancelled), covering roughly now ±8 h.
    /// Refreshed every <paramref name="refreshIntervalSeconds"/> seconds.
    /// </summary>
    public IReadOnlyList<ReminderMeeting> All { get; private set; } = Array.Empty<ReminderMeeting>();

    /// <summary>UTC timestamp of the last successful fetch, or <see cref="DateTime.MinValue"/> if never refreshed.</summary>
    public DateTime LastRefreshed { get; private set; } = DateTime.MinValue;

    /// <summary>Raised on the UI thread after each refresh attempt (successful or not).</summary>
    public event EventHandler? Refreshed;

    public MeetingCache(MeetingReminderService service, int refreshIntervalSeconds = 30)
    {
        _service = service;
        _timer = new System.Windows.Forms.Timer { Interval = refreshIntervalSeconds * 1000 };
        _timer.Tick += (_, _) => Refresh();
    }

    /// <summary>Performs an initial synchronous refresh then starts the background timer.</summary>
    public void Start()
    {
        Refresh();
        _timer.Start();
    }

    /// <summary>Immediately fetches fresh data from Outlook and raises <see cref="Refreshed"/>.</summary>
    public void Refresh()
    {
        var now = DateTime.Now;
        try
        {
            All = _service.FetchAll(now.Subtract(QueryHistoryWindow), now.Add(QueryFutureWindow));
            LastRefreshed = now;
        }
        catch
        {
            // Keep stale data on error; LastRefreshed not updated so callers can detect staleness.
        }
        Refreshed?.Invoke(this, EventArgs.Empty);
    }

    public void Dispose() => _timer.Dispose();
}
