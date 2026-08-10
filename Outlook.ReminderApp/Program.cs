using System.Runtime.Versioning;
using System.Threading;

namespace Outlook.ReminderApp;

[SupportedOSPlatform("windows")]
internal static class Program
{
    [STAThread]
    private static void Main()
    {
        using var mutex = new Mutex(true, "Outlook.ReminderApp.Singleton", out var isNewInstance);
        if (!isNewInstance)
        {
            return;
        }

        ApplicationConfiguration.Initialize();

        var uiContext = SynchronizationContext.Current ?? new WindowsFormsSynchronizationContext();

        using var reminderService = new MeetingReminderService();
        using var cache = new MeetingCache(reminderService, uiContext);
        using var syncScheduler = new SyncScheduler(uiContext);
        var syncStore = new SyncConfigStore();
        var syncUi = new SyncUiController(syncStore, syncScheduler);

        syncScheduler.SetRules(syncStore.Load());
        syncScheduler.Start();
        cache.Start();

        var agendaForm = new AgendaForm(reminderService, cache, syncUi);
        agendaForm.Show();
        Application.Run(new NotificationForm(reminderService, cache, syncUi));
    }
}
