using System.CommandLine;

namespace Outlook.Cli;

public static class SyncCommand
{
    public static Command Build()
    {
        var sourceOption = new Option<string>(
            name: "--source",
            description: "Source account name (sync FROM this calendar)")
        { IsRequired = true };

        var targetOption = new Option<string>(
            name: "--target",
            description: "Target account name (sync TO this calendar)")
        { IsRequired = true };

        var fromOption = new Option<string?>(
            name: "--from",
            description: "Start date (yyyy-MM-dd). Defaults to today.",
            getDefaultValue: () => null);

        var toOption = new Option<string?>(
            name: "--to",
            description: "End date (yyyy-MM-dd). Defaults to today + 90 days.",
            getDefaultValue: () => null);

        var modeOption = new Option<string>(
            name: "--mode",
            description: "Sync mode: 'block' (anonymous busy blocks) or 'copy' (copies title and description).",
            getDefaultValue: () => "block");

        var outsideHoursOption = new Option<bool>(
            name: "--outside-hours",
            description: "Only sync events that occur outside working hours (07:00–18:00).",
            getDefaultValue: () => false);

        var cmd = new Command("sync", "Sync events from one calendar to another")
        {
            sourceOption, targetOption, fromOption, toOption, modeOption, outsideHoursOption
        };

        cmd.SetHandler((string source, string target, string? from, string? to, string modeStr, bool outsideHours) =>
        {
            DateTime? fromDate = null;
            DateTime? toDate = null;

            if (from != null)
            {
                if (!DateTime.TryParseExact(from, "yyyy-MM-dd", null,
                        System.Globalization.DateTimeStyles.None, out var fd))
                {
                    Console.Error.WriteLine($"Invalid --from date '{from}'. Expected format: yyyy-MM-dd");
                    return;
                }
                fromDate = fd;
            }

            if (to != null)
            {
                if (!DateTime.TryParseExact(to, "yyyy-MM-dd", null,
                        System.Globalization.DateTimeStyles.None, out var td))
                {
                    Console.Error.WriteLine($"Invalid --to date '{to}'. Expected format: yyyy-MM-dd");
                    return;
                }
                toDate = td;
            }

            if (string.Equals(source, target, StringComparison.OrdinalIgnoreCase))
            {
                Console.Error.WriteLine("--source and --target must be different accounts.");
                return;
            }

            SyncMode mode;
            if (string.Equals(modeStr, "copy", StringComparison.OrdinalIgnoreCase))
                mode = SyncMode.Copy;
            else if (string.Equals(modeStr, "block", StringComparison.OrdinalIgnoreCase))
                mode = SyncMode.Block;
            else
            {
                Console.Error.WriteLine($"Invalid --mode '{modeStr}'. Expected 'block' or 'copy'.");
                return;
            }

            var effectiveFrom = (fromDate ?? DateTime.Today).Date;
            var effectiveTo = (toDate ?? DateTime.Today.AddDays(90)).Date;

            var modeLabel = mode == SyncMode.Copy ? "copy (title + description)" : "block (anonymous busy)";
            var filterLabel = outsideHours ? " · outside working hours only" : string.Empty;

            Console.WriteLine();
            Console.WriteLine($"  Source  : {source}");
            Console.WriteLine($"  Target  : {target}");
            Console.WriteLine($"  Range   : {effectiveFrom:yyyy-MM-dd} → {effectiveTo:yyyy-MM-dd}");
            Console.WriteLine($"  Mode    : {modeLabel}{filterLabel}");
            Console.WriteLine();

            var svc = new CalendarSyncService();
            var summary = svc.RunSync(source, target, effectiveFrom, effectiveTo, mode, outsideHours);

            Console.WriteLine($"  Created : {summary.Created}");
            Console.WriteLine($"  Deleted : {summary.Deleted}");
            Console.WriteLine($"  Unchanged: {summary.Skipped}");
            if (summary.Errors > 0)
                Console.WriteLine($"  Errors  : {summary.Errors}");
            Console.WriteLine();
            Console.WriteLine(summary.Created == 0 && summary.Deleted == 0
                ? "  Nothing to do — already in sync."
                : $"  Done. {summary.Created} created, {summary.Deleted} deleted.");
        }, sourceOption, targetOption, fromOption, toOption, modeOption, outsideHoursOption);

        return cmd;
    }
}
