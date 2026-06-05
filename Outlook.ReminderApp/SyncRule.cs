using Outlook.COM;

namespace Outlook.ReminderApp;

internal sealed class SyncRule
{
    public bool Enabled { get; set; } = true;
    public string SourceAccount { get; set; } = string.Empty;
    public string TargetAccount { get; set; } = string.Empty;
    public SyncMode Mode { get; set; } = SyncMode.Block;
    public bool OutsideWorkHoursOnly { get; set; }
    public int WorkDayStartHour { get; set; } = 7;
    public int WorkDayEndHour { get; set; } = 18;

    public SyncRule Clone()
    {
        return new SyncRule
        {
            Enabled = Enabled,
            SourceAccount = SourceAccount,
            TargetAccount = TargetAccount,
            Mode = Mode,
            OutsideWorkHoursOnly = OutsideWorkHoursOnly,
            WorkDayStartHour = WorkDayStartHour,
            WorkDayEndHour = WorkDayEndHour
        };
    }
}
