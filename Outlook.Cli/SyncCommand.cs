using System.CommandLine;
using Outlook.COM;

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
        cmd.Subcommands.Add(BuildPurge());

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

    private static Command BuildPurge()
    {
        var accountOpt = new Option<string>("--account") { Description = "Account to purge synced events from", Required = true };
        var fromOpt    = new Option<string?>("--from") { Description = "Start date (yyyy-MM-dd). Defaults to 2 years ago." };
        var toOpt      = new Option<string?>("--to")   { Description = "End date (yyyy-MM-dd). Defaults to 2 years from now." };

        var cmd = new Command("purge", "Delete all synced/blocked events created by outlook sync from a calendar");
        cmd.Options.Add(accountOpt);
        cmd.Options.Add(fromOpt);
        cmd.Options.Add(toOpt);
        cmd.SetAction(ctx =>
        {
            var account = ctx.GetValue(accountOpt)!;
            var fromStr = ctx.GetValue(fromOpt);
            var toStr   = ctx.GetValue(toOpt);

            var from = fromStr != null
                ? DateTime.ParseExact(fromStr, "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture)
                : DateTime.Today.AddYears(-2);
            var to = toStr != null
                ? DateTime.ParseExact(toStr, "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture)
                : DateTime.Today.AddYears(2);

            using var svc = new OutlookCalendarService();
            var events = svc.ListEvents(from, to, account, bodyLength: int.MaxValue);
            var synced = events
                .Where(e => e.TryGetValue("body", out var b) && b is string s && s.Contains("[outlook-sync:"))
                .ToList();

            Console.WriteLine();
            Console.WriteLine($"  Account : {account}");
            Console.WriteLine($"  Range   : {from:yyyy-MM-dd} to {to:yyyy-MM-dd}");
            Console.WriteLine($"  Found   : {synced.Count} synced event(s) to delete");
            Console.WriteLine();

            int deleted = 0, errors = 0;
            foreach (var ev in synced)
            {
                var id = (string?)ev["id"];
                if (id == null) continue;
                try
                {
                    svc.DeleteEvent(id, account);
                    deleted++;
                }
                catch
                {
                    errors++;
                }
            }

            Console.WriteLine(deleted == 0
                ? "  Nothing to purge -- no synced events found."
                : $"  Done. {deleted} deleted" + (errors > 0 ? $", {errors} errors." : "."));
        });
        return cmd;
    }
}



