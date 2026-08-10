namespace Outlook.ReminderApp;

/// <summary>
/// Maintains a periodically refreshed in-memory snapshot of meetings fetched from Outlook.
/// The Outlook COM call runs on a dedicated background STA thread so the UI thread is never
/// blocked. Results are marshalled back to the UI thread before updating state and raising events.
/// </summary>
internal sealed class MeetingCache : IDisposable
{
    private static readonly TimeSpan QueryHistoryWindow = TimeSpan.FromHours(8);
    private static readonly TimeSpan QueryFutureWindow  = TimeSpan.FromHours(8);

    private readonly MeetingReminderService _service;
    private readonly SynchronizationContext _uiContext;
    private readonly System.Windows.Forms.Timer _timer;
    private readonly int _refreshIntervalMs;
    private int _isRefreshing; // 0 = idle, 1 = in progress; accessed via Interlocked

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

    public MeetingCache(MeetingReminderService service, SynchronizationContext uiContext, int refreshIntervalSeconds = 30)
    {
        _service = service;
        _uiContext = uiContext;
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

    /// <summary>
    /// Triggers a refresh on a background STA thread so the UI message loop stays responsive
    /// even if Outlook is slow or temporarily unresponsive. Skipped if a refresh is already running.
    /// </summary>
    public void Refresh()
    {
        if (Interlocked.CompareExchange(ref _isRefreshing, 1, 0) != 0)
            return;

        var now = DateTime.Now;
        var thread = new Thread(() =>
        {
            IReadOnlyList<ReminderMeeting>? result = null;
            try
            {
                result = _service.FetchAll(now.Subtract(QueryHistoryWindow), now.Add(QueryFutureWindow));
            }
            catch { }

            _uiContext.Post(_ =>
            {
                if (result is not null)
                {
                    All = result;
                    LastRefreshed = now;
                }
                Interlocked.Exchange(ref _isRefreshing, 0);
                Refreshed?.Invoke(this, EventArgs.Empty);
            }, null);
        });
        thread.IsBackground = true;
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
    }

    public void Dispose() => _timer.Dispose();
}
