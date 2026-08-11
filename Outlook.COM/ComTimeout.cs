using System.Collections.Concurrent;
using System.Runtime.ExceptionServices;
using System.Runtime.Versioning;

namespace Outlook.COM;

/// <summary>
/// Runs Outlook COM work on a single dedicated STA worker thread that lives for the process
/// lifetime, enforcing a hard timeout on each call.
/// </summary>
/// <remarks>
/// <para>
/// Previous design spawned a fresh STA thread per call. On timeout the thread was abandoned
/// while still holding the <c>Outlook.Application</c> RCW (the <c>using</c> never ran because the
/// lambda never returned). Each leaked RCW kept Outlook's main STA thread busy, making subsequent
/// calls slower and more likely to time out — a death spiral.
/// </para>
/// <para>
/// This design uses one persistent STA worker thread with a <see cref="BlockingCollection{T}"/>
/// queue. No new thread or RCW is created per call. On timeout the caller throws, but the worker
/// keeps processing the in-flight call; when it eventually completes the <c>using</c>/Dispose
/// inside the lambda runs and releases the RCW. No leak. Subsequent calls simply queue behind the
/// in-flight one (contention, not a leak).
/// </para>
/// <para>
/// <c>outlook-cli</c> and <c>Outlook.ReminderApp</c> are separate processes, each with their own
/// worker thread and their own cached <c>Outlook.Application</c> RCW — but both RCWs attach to the
/// same single-instance <c>OUTLOOK.EXE</c>, whose main thread only runs one call at a time. A named
/// <see cref="Mutex"/> (<see cref="CrossProcessGateName"/>) gates actual execution so the two
/// processes take turns instead of racing. Waiting for that gate and running the call itself are
/// budgeted separately (<see cref="GateWaitLimit"/> then <see cref="CallLimit"/>, ~60s worst case)
/// so a call that is merely queued behind the other process's in-flight call is never confused with
/// one where Outlook itself is hung.
/// </para>
/// </remarks>
[SupportedOSPlatform("windows")]
internal static class ComTimeout
{
    private const string CrossProcessGateName = "OutlookMcp_ComAccessGate";

    internal static readonly TimeSpan GateWaitLimit = TimeSpan.FromSeconds(30);
    internal static readonly TimeSpan CallLimit = TimeSpan.FromSeconds(30);

    private static readonly Mutex _gate = new(initiallyOwned: false, CrossProcessGateName);
    private static readonly BlockingCollection<WorkItem> _queue = new();
    private static readonly Thread _worker;
    private static readonly int _workerThreadId;

    static ComTimeout()
    {
        _worker = new Thread(WorkerLoop)
        {
            IsBackground = true,
            Name = "Outlook-COM-Worker"
        };
        _worker.SetApartmentState(ApartmentState.STA);
        _workerThreadId = _worker.ManagedThreadId;
        _worker.Start();
    }

    private static void WorkerLoop()
    {
        foreach (var item in _queue.GetConsumingEnumerable())
        {
            // Execute catches all exceptions internally; Done.Set always fires.
            item.Execute();
        }
    }

    /// <summary>
    /// Enqueues <paramref name="work"/> to the persistent STA worker and waits up to
    /// <see cref="GateWaitLimit"/> + <see cref="CallLimit"/> combined. Throws
    /// <see cref="TimeoutException"/> on timeout — the worker thread keeps running the call (or
    /// keeps waiting for the cross-process gate) so no RCW is leaked and the other process is never
    /// blocked forever. If already on the worker thread the work runs inline to avoid a
    /// self-deadlock — this also means it piggybacks on the gate the outer call already holds,
    /// with no double-acquire.
    /// </summary>
    internal static T Run<T>(Func<T> work)
    {
        if (Thread.CurrentThread.ManagedThreadId == _workerThreadId)
            return work();

        var item = WorkItem<T>.Create(work);
        _queue.Add(item);

        var totalWait = GateWaitLimit + CallLimit;
        if (!item.Done.Wait(totalWait))
            throw new TimeoutException(
                $"Outlook COM operation did not complete within {totalWait.TotalSeconds:0} seconds.");

        return item.GetResult();
    }

    internal static void Run(Action work) =>
        Run<object?>(() => { work(); return null; });

    private abstract class WorkItem
    {
        public abstract void Execute();
    }

    private sealed class WorkItem<T> : WorkItem
    {
        private readonly Func<T> _work;
        private T? _result;
        private Exception? _error;
        public readonly ManualResetEventSlim Done = new();

        private WorkItem(Func<T> work) => _work = work;

        public static WorkItem<T> Create(Func<T> work) => new(work);

        public override void Execute()
        {
            try
            {
                bool acquired;
                try
                {
                    acquired = _gate.WaitOne(GateWaitLimit);
                }
                catch (AbandonedMutexException)
                {
                    // Previous owner's process died mid-call; the gate is ours now.
                    acquired = true;
                }

                if (!acquired)
                {
                    _error = new TimeoutException(
                        "Outlook is busy handling a request from another Outlook process " +
                        $"(outlook-cli / Outlook.ReminderApp). Timed out after {GateWaitLimit.TotalSeconds:0}s " +
                        "waiting for exclusive access.");
                    return;
                }

                try { _result = _work(); }
                catch (Exception ex) { _error = ex; }
                finally { _gate.ReleaseMutex(); }
            }
            finally { Done.Set(); }
        }

        public T GetResult()
        {
            if (_error is not null)
                ExceptionDispatchInfo.Capture(_error).Throw();
            return _result!;
        }
    }
}
