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
    private static dynamic? _app;
    private static readonly object _lock = new();

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
                var type = Type.GetTypeFromProgID("Outlook.Application")
                    ?? throw new InvalidOperationException(
                        "Microsoft Outlook is not installed or not registered on this system.");

                return Activator.CreateInstance(type)
                    ?? throw new InvalidOperationException("Failed to create Outlook.Application instance.");
            });

            return _app;
        }
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
        }

        if (old is not null)
        {
            // Release on the STA worker thread — the RCW was created there.
            try { ComTimeout.Run(() => { try { Marshal.ReleaseComObject(old); } catch { } }); }
            catch { /* timeout or worker error — RCW will be collected by GC */ }
        }
    }

}