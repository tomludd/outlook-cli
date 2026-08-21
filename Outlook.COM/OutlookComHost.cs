using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace Outlook.COM;

/// <summary>
/// Owns the single shared <c>Outlook.Application</c> RCW for the process.
/// Created lazily on the persistent STA worker thread (via <see cref="ComTimeout"/>)
/// and never explicitly released — the RCW dies with the process when the STA
/// worker (a background thread) exits.
/// </summary>
/// <remarks>
/// <para>
/// All three services (<see cref="OutlookCalendarService"/>,
/// <see cref="OutlookContactService"/>, <see cref="OutlookMailService"/>)
/// delegate <c>GetOutlookApp()</c> here so they share one RCW instead of each
/// creating and destroying their own. This eliminates per-call ROT lookups,
/// avoids the timeout race where <c>Dispose</c> releases an RCW the worker is
/// still using, and ensures the RCW is created and used on the same STA thread.
/// </para>
/// <para>
/// The RCW is never released explicitly. Releasing a shared RCW is dangerous
/// (one service's release would break another's in-flight call) and unnecessary
/// (the STA worker is a background thread; the RCW is freed when the process
/// exits). For short-lived processes like <c>outlook-cli</c> the RCW lives for
/// the few seconds the CLI runs — negligible.
/// </para>/// <para>
/// If Outlook is closed and reopened while the process is running, the cached
/// RCW becomes stale. <see cref="OutlookComInvoker"/> automatically detects
/// RPC/server-unavailable errors, calls <see cref="Reset"/>, and retries once
/// so callers reconnect to the new Outlook instance without manual handling.
/// </para>/// </remarks>
[SupportedOSPlatform("windows")]
internal static class OutlookComHost
{
    private const string OutlookProcessName = "OUTLOOK";

    private static dynamic? _app;
    private static readonly object _lock = new();
    private static int? _spawnedPid;
    private static DateTime? _spawnedStartTime;

    /// <summary>
    /// Returns the process-wide <c>Outlook.Application</c> RCW, creating it on
    /// the STA worker thread if it does not yet exist. Thread-safe.
    /// </summary>
    internal static dynamic GetApp()
    {
        if (_app is not null)
            return _app;

        lock (_lock)
        {
            if (_app is not null)
                return _app;

            // Create on the STA worker so the RCW is marshalled to the same
            // thread that will invoke all COM calls through ComTimeout.Run.
            _app = ComTimeout.Run(() =>
            {
                var pidsBefore = GetRunningOutlookPids();

                var type = Type.GetTypeFromProgID("Outlook.Application")
                    ?? throw new InvalidOperationException(
                        "Microsoft Outlook is not installed or not registered on this system.");

                var instance = Activator.CreateInstance(type)
                    ?? throw new InvalidOperationException("Failed to create Outlook.Application instance.");

                TrackIfWeSpawnedOutlook(pidsBefore);
                return instance;
            });

            return _app;
        }
    }

    private static HashSet<int> GetRunningOutlookPids()
    {
        var pids = new HashSet<int>();
        foreach (var process in Process.GetProcessesByName(OutlookProcessName))
        {
            using (process) pids.Add(process.Id);
        }
        return pids;
    }

    /// <summary>
    /// Records the PID of the OUTLOOK.EXE process that came into existence as a side effect of
    /// the <c>Activator.CreateInstance</c> call above — i.e. one that did not exist before we
    /// asked COM to activate <c>Outlook.Application</c>, meaning COM launched it in the
    /// background rather than binding to an instance the user already had open. Only a process
    /// we can prove we spawned this way is ever eligible for <see cref="OutlookHealthMonitor"/>
    /// to kill; an instance the user started is never tracked here and so can never be killed.
    /// </summary>
    private static void TrackIfWeSpawnedOutlook(HashSet<int> pidsBefore)
    {
        foreach (var process in Process.GetProcessesByName(OutlookProcessName))
        {
            using (process)
            {
                if (!pidsBefore.Contains(process.Id))
                {
                    _spawnedPid = process.Id;
                    _spawnedStartTime = process.StartTime;
                    return;
                }
            }
        }
    }

    /// <summary>
    /// Returns the OUTLOOK.EXE process we spawned in the background, if it is still running
    /// under the same PID and start time we recorded (guards against PID reuse). Returns null
    /// if we never spawned one (the user's own Outlook was already running) or it has since exited.
    /// Caller owns the returned <see cref="Process"/> and must dispose it.
    /// </summary>
    internal static Process? TryGetSpawnedProcess()
    {
        if (_spawnedPid is not int pid)
            return null;

        try
        {
            var process = Process.GetProcessById(pid);
            if (process.StartTime == _spawnedStartTime)
                return process;

            process.Dispose();
        }
        catch (ArgumentException)
        {
            // Process no longer exists.
        }

        return null;
    }

    /// <summary>
    /// Safely releases a COM object, swallowing any failure (e.g. already
    /// released, or the underlying object is gone). Use in <c>finally</c>
    /// blocks for intermediate COM objects (ns, stores, folder, items, etc.)
    /// so they don't pile up waiting for GC finalization.
    /// </summary>
    internal static void Release(object? obj)
    {
        if (obj is null) return;
        try { Marshal.ReleaseComObject(obj); }
        catch { /* already released or invalid */ }
    }

    /// <summary>
    /// Clears the cached RCW so the next <see cref="GetApp"/> reconnects to
    /// Outlook. Call this when a COM call fails with an RPC/server-unavailable
    /// error (Outlook was closed/restarted). The old RCW is released on the
    /// STA worker thread to avoid cross-thread marshalling issues.
    /// </summary>
    internal static void Reset()
    {
        dynamic? old;
        lock (_lock)
        {
            old = _app;
            _app = null;
            _spawnedPid = null;
            _spawnedStartTime = null;
        }

        if (old is not null)
        {
            // Release on the STA worker thread — the RCW was created there.
            try { ComTimeout.Run(() => { try { Marshal.ReleaseComObject(old); } catch { } }); }
            catch { /* timeout or worker error — RCW will be collected by GC */ }
        }
    }

}