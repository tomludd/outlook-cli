namespace Outlook.ReminderApp;

internal sealed class SyncUiController
{
    private readonly SyncConfigStore _store;
    private readonly SyncScheduler _scheduler;

    public SyncUiController(SyncConfigStore store, SyncScheduler scheduler)
    {
        _store = store;
        _scheduler = scheduler;
    }

    public void ShowConfig(IWin32Window? owner)
    {
        var rules = _store.Load();
        using var form = new SyncConfigForm(rules, _scheduler);
        var result = owner is null ? form.ShowDialog() : form.ShowDialog(owner);
        if (result != DialogResult.OK)
        {
            return;
        }

        var updatedRules = form.Rules.ToList();
        _store.Save(updatedRules);
        _scheduler.SetRules(updatedRules);
    }
}
