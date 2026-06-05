using System.Diagnostics;
using System.Threading;
using Outlook.COM;

namespace Outlook.ReminderApp;

internal sealed class SyncScheduler : IDisposable
{
    private static readonly int[] ScheduleMinutes = { 12, 27, 42, 57 };
    private static readonly TimeSpan DefaultRange = TimeSpan.FromDays(90);

    private readonly System.Windows.Forms.Timer _timer;
    private readonly SynchronizationContext _uiContext;
    private IReadOnlyList<SyncRule> _rules = Array.Empty<SyncRule>();
    private bool _isRunning;
    private readonly List<SyncRunLogEntry> _log = new();
    private readonly object _logLock = new();
    private DateTime? _lastRunAt;
    private DateTime? _lastErrorAt;
    private string? _lastError;

    public event EventHandler? LogsUpdated;

    public SyncScheduler(SynchronizationContext uiContext)
    {
        _uiContext = uiContext;
        _timer = new System.Windows.Forms.Timer();
        _timer.Tick += (_, _) => OnTick();
    }

    public void Start()
    {
        ScheduleNext(DateTime.Now);
        _timer.Start();
    }

    public void Stop()
    {
        _timer.Stop();
    }

    public void SetRules(IReadOnlyList<SyncRule> rules)
    {
        _rules = rules ?? Array.Empty<SyncRule>();
    }

    public SyncStatusSnapshot GetStatus()
    {
        return new SyncStatusSnapshot(_lastRunAt, _lastErrorAt, _lastError);
    }

    public IReadOnlyList<SyncRunLogEntry> GetRecentLogs(int max = 50)
    {
        lock (_logLock)
        {
            if (_log.Count <= max) return _log.ToList();
            return _log.Skip(Math.Max(0, _log.Count - max)).ToList();
        }
    }

    public void RunAllNow()
    {
        if (_isRunning) return;
        _isRunning = true;
        RunSyncsOnStaThread(DateTime.Now);
    }

    private void OnTick()
    {
        ScheduleNext(DateTime.Now);
        if (_isRunning) return;
        _isRunning = true;
        RunSyncsOnStaThread(DateTime.Now);
    }

    private void RunSyncsOnStaThread(DateTime now)
    {
        var rulesSnapshot = _rules.ToList();
        var thread = new Thread(() => RunSyncs(rulesSnapshot, now))
        {
            IsBackground = true
        };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
    }

    private void RunSyncs(IReadOnlyList<SyncRule> rules, DateTime now)
    {
        try
        {
            _lastRunAt = now;
            var from = now.Date;
            var to = from.Add(DefaultRange);
            var anyRule = false;

            foreach (var rule in rules)
            {
                if (!rule.Enabled) continue;
                if (string.IsNullOrWhiteSpace(rule.SourceAccount)) continue;
                if (string.IsNullOrWhiteSpace(rule.TargetAccount)) continue;
                anyRule = true;

                try
                {
                    var svc = new CalendarSyncService();
                    var summary = svc.RunSync(
                        rule.SourceAccount,
                        rule.TargetAccount,
                        from,
                        to,
                        rule.Mode,
                        rule.OutsideWorkHoursOnly,
                        rule.WorkDayStartHour,
                        rule.WorkDayEndHour);

                    AddLog(new SyncRunLogEntry
                    {
                        Timestamp = DateTime.Now,
                        RuleLabel = BuildRuleLabel(rule),
                        Status = "OK",
                        Message = $"{summary.Created} created, {summary.Deleted} deleted, {summary.Skipped} skipped"
                    });
                }
                catch (Exception ex)
                {
                    _lastErrorAt = DateTime.Now;
                    _lastError = ex.Message;
                    Debug.WriteLine($"Sync rule failed: {ex.Message}");
                    AddLog(new SyncRunLogEntry
                    {
                        Timestamp = DateTime.Now,
                        RuleLabel = BuildRuleLabel(rule),
                        Status = "ERROR",
                        Message = ex.Message
                    });
                }
            }

            if (!anyRule)
            {
                AddLog(new SyncRunLogEntry
                {
                    Timestamp = DateTime.Now,
                    Status = "SKIP",
                    Message = "No enabled rules to run"
                });
            }
        }
        finally
        {
            _uiContext.Post(_ => _isRunning = false, null);
        }
    }

    private void AddLog(SyncRunLogEntry entry)
    {
        lock (_logLock)
        {
            _log.Add(entry);
            var cutoff = DateTime.Now.AddHours(-24);
            _log.RemoveAll(x => x.Timestamp < cutoff);
            if (_log.Count > 200)
            {
                _log.RemoveRange(0, _log.Count - 200);
            }
        }
        _uiContext.Post(_ => LogsUpdated?.Invoke(this, EventArgs.Empty), null);
    }

    private static string BuildRuleLabel(SyncRule rule)
    {
        var mode = rule.Mode == SyncMode.Copy ? "copy" : "block";
          return $"{rule.SourceAccount} -> {rule.TargetAccount} ({mode})";
    }

    private void ScheduleNext(DateTime now)
    {
        var next = ComputeNextRun(now);
        var delay = next - now;
        var ms = (int)Math.Max(1000, Math.Min(delay.TotalMilliseconds, int.MaxValue));
        _timer.Interval = ms;
    }

    private static DateTime ComputeNextRun(DateTime now)
    {
        foreach (var minute in ScheduleMinutes)
        {
            var candidate = new DateTime(now.Year, now.Month, now.Day, now.Hour, minute, 0, now.Kind);
            if (candidate > now)
            {
                return candidate;
            }
        }

        var nextHour = now.AddHours(1);
        return new DateTime(nextHour.Year, nextHour.Month, nextHour.Day, nextHour.Hour, ScheduleMinutes[0], 0, now.Kind);
    }

    public void Dispose()
    {
        _timer.Stop();
        _timer.Dispose();
    }
}

internal sealed record SyncStatusSnapshot(DateTime? LastRunAt, DateTime? LastErrorAt, string? LastError);
