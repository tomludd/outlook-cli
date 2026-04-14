using System.CommandLine;

namespace Outlook.Cli;

public static class SyncCommand
{
    public static Command Build()
    {
        var sourceOption = new Option<string>("--source") { Description = "Source account name (sync FROM this calendar)", Required = true };
        var targetOption = new Option<string>("--target") { Description = "Target account name (sync TO this calendar)", Required = true };
        var fromOption   = new Option<string?>("--from") { Description = "Start date (yyyy-MM-dd). Defaults to today." };
        var toOption     = new Option<string?>("--to") { Description = "End date (yyyy-MM-dd). Defaults to today + 90 days." };
        var modeOption   = new Option<string>("--mode") { Description = "Sync mode: 'block' (anonymous busy blocks) or 'copy' (copies title and description).", DefaultValueFactory = _ => "block" };
        var outsideHoursOption = new Option<bool>("--outside-hours") { Description = "Only sync events outside working hours (07:00-18:00).", DefaultValueFactory = _ => false };

        var cmd = new Command("sync", "Sync events from one calendar to another");
        cmd.Options.Add(sourceOption);
        cmd.Options.Add(targetOption);
        cmd.Options.Add(fromOption);
        cmd.Options.Add(toOption);
        cmd.Options.Add(modeOption);
        cmd.Options.Add(outsideHoursOption);

        cmd.SetAction(ctx =>
        {
            var source       = ctx.GetValue(sourceOption)!;
            var target       = ctx.GetValue(targetOption)!;
            var from         = ctx.GetValue(fromOption);
            var to           = ctx.GetValue(toOption);
            var modeStr      = ctx.GetValue(modeOption)!;
            var outsideHours = ctx.GetValue(outsideHoursOption);

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
            var effectiveTo   = (toDate ?? DateTime.Today.AddDays(90)).Date;

            var modeLabel   = mode == SyncMode.Copy ? "copy (title + description)" : "block (anonymous busy)";
            var filterLabel = outsideHours ? " * outside working hours only" : string.Empty;

            Console.WriteLine();
            Console.WriteLine($"  Source  : {source}");
            Console.WriteLine($"  Target  : {target}");
            Console.WriteLine($"  Range   : {effectiveFrom:yyyy-MM-dd} to {effectiveTo:yyyy-MM-dd}");
            Console.WriteLine($"  Mode    : {modeLabel}{filterLabel}");
            Console.WriteLine();

            var svc     = new CalendarSyncService();
            var summary = svc.RunSync(source, target, effectiveFrom, effectiveTo, mode, outsideHours);

            Console.WriteLine($"  Created  : {summary.Created}");
            Console.WriteLine($"  Deleted  : {summary.Deleted}");
            Console.WriteLine($"  Unchanged: {summary.Skipped}");
            if (summary.Errors > 0)
                Console.WriteLine($"  Errors   : {summary.Errors}");
            Console.WriteLine();
            Console.WriteLine(summary.Created == 0 && summary.Deleted == 0
                ? "  Nothing to do -- already in sync."
                : $"  Done. {summary.Created} created, {summary.Deleted} deleted.");
        });

        return cmd;
    }
}



