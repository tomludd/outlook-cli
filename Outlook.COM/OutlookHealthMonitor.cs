using System.Runtime.Versioning;
using System.Threading;

namespace Outlook.COM;

/// <summary>
/// Tracks whether Outlook appears to be hung and, if so, whether it is safe for a caller to
/// restart it. "Safe" here means the running <c>OUTLOOK.EXE</c> is one <see cref="OutlookComHost"/>
/// itself launched in the background (e.g. a scheduled sync or the reminder app's poll started it
/// because it wasn't already running) — never an instance the user opened themselves. Because
/// Outlook enforces a single instance, a hang in a background-launched Outlook blocks the user's
/// own UI too (they end up looking at the same wedged process), but this class deliberately can't
/// tell the two apart once a user starts interacting with a process we spawned — so callers should
/// still confirm with the user before actually killing anything.
/// </summary>
[SupportedOSPlatform("windows")]
public static class OutlookHealthMonitor
{
    /// <summary>
    /// Number of consecutive full <see cref="ComTimeout"/> timeouts (Outlook never responded,
    /// as opposed to a transient "busy" rejection) before <see cref="IsLikelyHung"/> reports true.
    /// </summary>
    private const int HungThreshold = 2;

    private static int _consecutiveTimeouts;

    /// <summary>Number of consecutive COM calls that have failed with a hard timeout.</summary>
    public static int ConsecutiveTimeouts => _consecutiveTimeouts;

    /// <summary>
    /// True once Outlook has failed to respond at all for <see cref="HungThreshold"/> calls in a
    /// row <em>and</em> the running Outlook is one we launched ourselves in the background
    /// (see <see cref="HasRestartableProcess"/>). Never true for an Outlook instance the user
    /// started — there is nothing for a caller to safely restart in that case.
    /// </summary>
    public static bool IsLikelyHung => _consecutiveTimeouts >= HungThreshold && HasRestartableProcess;

    /// <summary>
    /// True if we're holding onto the PID of an OUTLOOK.EXE process we spawned in the background
    /// and it is still running.
    /// </summary>
    public static bool HasRestartableProcess
    {
        get
        {
            using var process = OutlookComHost.TryGetSpawnedProcess();
            return process is not null;
        }
    }

    internal static void RecordTimeout() => Interlocked.Increment(ref _consecutiveTimeouts);

    internal static void RecordSuccess() => Interlocked.Exchange(ref _consecutiveTimeouts, 0);

    /// <summary>
    /// Kills the background Outlook process we spawned ourselves and resets the cached RCW so the
    /// next call launches a fresh instance. Does nothing — and returns false — if Outlook was never
    /// launched by us in the background (i.e. the running instance is the user's own). Callers
    /// should confirm with the user first: this forcibly terminates OUTLOOK.EXE, which can drop
    /// unsaved work in any window open against that process.
    /// </summary>
    public static bool TryRestartSpawnedOutlook()
    {
        using var process = OutlookComHost.TryGetSpawnedProcess();
        if (process is null)
            return false;

        try
        {
            process.Kill();
            process.WaitForExit(5000);
        }
        catch
        {
            return false;
        }
        finally
        {
            OutlookComHost.Reset();
            Interlocked.Exchange(ref _consecutiveTimeouts, 0);
        }

        return true;
    }
}
