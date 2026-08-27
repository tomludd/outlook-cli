using Outlook.COM;

namespace Outlook.ReminderApp;

/// <summary>
/// Maintains a periodically refreshed in-memory snapshot of meetings fetched from Outlook.
/// The Outlook COM call runs on a dedicated background STA thread so the UI thread is never
/// blocked. Results are marshalled back to the UI thread before updating state and raising events.
/// </summary>
internal sealed class MeetingCache : IDisposable
{
    /// <summary>Number of days before today to fetch (day-aligned).</summary>
    private const int HistoryDays = 7;

    /// <summary>Number of days after today to fetch (day-aligned, inclusive).</summary>
    private const int FutureDays = 7;

    private readonly MeetingReminderService _service;
    private readonly SynchronizationContext _uiContext;
    private readonly System.Windows.Forms.Timer _timer;
    private readonly int _refreshIntervalMs;
    private int _isRefreshing; // 0 = idle, 1 = in progress; accessed via Interlocked

    /// <summary>
    /// All fetched meetings (including cancelled), covering today ±7 calendar days.
    /// Refreshed every <paramref name="refreshIntervalSeconds"/> seconds.
    /// </summary>
    public IReadOnlyList<ReminderMeeting> All { get; private set; } = Array.Empty<ReminderMeeting>();

    /// <summary>UTC timestamp of the last successful fetch, or <see cref="DateTime.MinValue"/> if never refreshed.</summary>
    public DateTime LastRefreshed { get; private set; } = DateTime.MinValue;

    /// <summary>True when the most recent refresh attempt failed (Outlook returned no data).</summary>
    public bool LastRefreshFailed { get; private set; }

    /// <summary>Error message from the most recent failed refresh, or null if none.</summary>
    public string? LastError { get; private set; }

    /// <summary>
    /// True when Outlook has stopped responding to COM calls entirely (not merely rejected one as
    /// busy) for long enough that it looks hung, and the process is one we launched in the
    /// background ourselves rather than one the user started. See <see cref="OutlookHealthMonitor"/>.
    /// </summary>
    public bool IsOutlookLikelyHung => OutlookHealthMonitor.IsLikelyHung;

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
    /// Triggers a refresh on a background thread so the UI message loop stays responsive
    /// even if Outlook is slow or temporarily unresponsive. Skipped if a refresh is already running.
    /// Uses Task.Run (MTA thread pool) rather than a raw STA thread — ComTimeout.Run already
    /// creates its own STA thread for the COM work, and nesting two STA blocking waits deadlocks.
    /// </summary>
    public void Refresh()
    {
        if (Interlocked.CompareExchange(ref _isRefreshing, 1, 0) != 0)
            return;

        var now = DateTime.Now;
        var from = now.Date.AddDays(-HistoryDays);
        var to   = now.Date.AddDays(FutureDays + 1); // +1 to include the full last day
        Task.Run(() =>
        {
            IReadOnlyList<ReminderMeeting>? result = null;
            try
            {
                result = _service.FetchAll(from, to);
            }
            catch (Exception ex)
            {
                LastError = ex.Message;
            }

            _uiContext.Post(_ =>
            {
                if (result is not null)
                {
                    All = result;
                    LastRefreshed = now;
                    LastRefreshFailed = false;
                    LastError = null;
                }
                else
                {
                    LastRefreshFailed = true;
                }
                Interlocked.Exchange(ref _isRefreshing, 0);
                Refreshed?.Invoke(this, EventArgs.Empty);
            }, null);
        });
    }

    public void Dispose() => _timer.Dispose();
}
