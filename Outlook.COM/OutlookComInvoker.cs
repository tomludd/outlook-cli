using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace Outlook.COM;

/// <summary>
/// Exception result codes that indicate the cached Outlook.Application RCW
/// has become stale (Outlook was closed or restarted).
/// </summary>
internal static class RpcHResult
{
    /// <summary>RPC server unavailable — Outlook was shut down.</summary>
    internal const int ServerUnavailable = unchecked((int)0x800706BA);

    /// <summary>Call was canceled by the message filter — Outlook is no longer responding.</summary>
    internal const int CallCanceled = unchecked((int)0x80010002);

    /// <summary>Call was rejected by the callee — Outlook likely busy or shutting down.</summary>
    internal const int CallRejected = unchecked((int)0x80010001);

    /// <summary>The object invoked has disconnected from its clients.</summary>
    internal const int ServerDisconnected = unchecked((int)0x80010108);
}

/// <summary>
/// Single entry point for Outlook COM calls in this library.
/// Combines the STA-worker timeout enforcement from <see cref="ComTimeout"/>
/// with automatic recovery from two distinct COM failure modes: a stale
/// <c>Outlook.Application</c> RCW (Outlook was closed/restarted) and a
/// transient "server busy" rejection (Outlook was mid-call).
/// </summary>
/// <remarks>
/// Use this instead of calling <see cref="ComTimeout.Run"/> or
/// <see cref="OutlookComHost"/> directly from service code.
/// </remarks>
[SupportedOSPlatform("windows")]
internal static class OutlookComInvoker
{
    private const int MaxBusyAttempts = 3;
    private static readonly TimeSpan InitialBusyDelay = TimeSpan.FromMilliseconds(250);

    /// <summary>
    /// Runs <paramref name="work"/> on the STA worker. Retries once, after resetting the
    /// cached RCW, if the failure looks like a stale Outlook connection. Retries a few times
    /// with a short backoff — without touching the RCW, which is still perfectly valid — if
    /// Outlook merely rejected the call because it was busy with something else.
    /// </summary>
    internal static T Run<T>(Func<T> work)
    {
        var busyDelay = InitialBusyDelay;

        for (var attempt = 1; ; attempt++)
        {
            try
            {
                var result = ComTimeout.Run(work);
                OutlookHealthMonitor.RecordSuccess();
                return result;
            }
            catch (TimeoutException)
            {
                // Outlook never responded at all — as opposed to rejecting the call because it
                // was busy — feeds OutlookHealthMonitor's hang detection.
                OutlookHealthMonitor.RecordTimeout();
                throw;
            }
            catch (COMException ex) when (IsStaleConnection(ex))
            {
                OutlookComHost.Reset();
                var result = ComTimeout.Run(work);
                OutlookHealthMonitor.RecordSuccess();
                return result;
            }
            catch (COMException ex) when (IsBusyRejected(ex) && attempt < MaxBusyAttempts)
            {
                Thread.Sleep(busyDelay);
                busyDelay += busyDelay;
            }
        }
    }

    /// <summary>
    /// Runs <paramref name="work"/> on the STA worker with the same stale-connection and
    /// busy-retry handling as <see cref="Run{T}"/>.
    /// </summary>
    internal static void Run(Action work) =>
        Run<object?>(() => { work(); return null; });

    /// <summary>The cached RCW itself is dead — Outlook was closed or restarted.</summary>
    private static bool IsStaleConnection(COMException ex)
    {
        return ex.HResult == RpcHResult.ServerUnavailable
            || ex.HResult == RpcHResult.ServerDisconnected;
    }

    /// <summary>The RCW is fine — Outlook was just mid-call and rejected this one.</summary>
    private static bool IsBusyRejected(COMException ex)
    {
        return ex.HResult == RpcHResult.CallRejected
            || ex.HResult == RpcHResult.CallCanceled;
    }
}
