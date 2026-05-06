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
    private readonly int _refreshIntervalMs;

    /// <summary>
    /// All fetched meetings (including cancelled), covering roughly now ±8 h.
    /// Refreshed every <paramref name="refreshIntervalSeconds"/> seconds.
    /// </summary>
    public IReadOnlyList<ReminderMeeting> All { get; private set; } = Array.Empty<ReminderMeeting>();

    /// <summary>UTC timestamp of the last successful fetch, or <see cref="DateTime.MinValue"/> if never refreshed.</summary>
    public DateTime LastRefreshed { get; private set; } = DateTime.MinValue;

    /// <summary>True once the first refresh has completed successfully.</summary>
    public bool IsLoaded => LastRefreshed > DateTime.MinValue;

    /// <summary>Raised on the UI thread after each refresh attempt (successful or not).</summary>
    public event EventHandler? Refreshed;

    public MeetingCache(MeetingReminderService service, int refreshIntervalSeconds = 30)
    {
        _service = service;
        _refreshIntervalMs = refreshIntervalSeconds * 1000;
        _timer = new System.Windows.Forms.Timer { Interval = _refreshIntervalMs };
        _timer.Tick += (_, _) => Refresh();
    }

    /// <summary>
    /// Schedules the first refresh to run shortly after the message loop starts,
    /// keeping the UI thread free during startup so the window and taskbar icon appear immediately.
    /// </summary>
    public void Start()
    {
        var startupTimer = new System.Windows.Forms.Timer { Interval = 100 };
        startupTimer.Tick += (_, _) =>
        {
            startupTimer.Stop();
            startupTimer.Dispose();
            Refresh();
            _timer.Start();
        };
        startupTimer.Start();
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
