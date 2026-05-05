using System.Runtime.Versioning;

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

        using var reminderService = new MeetingReminderService();
        Application.Run(new MainForm(reminderService));
    }
}
