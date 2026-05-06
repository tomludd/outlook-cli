using System.CommandLine;
using Outlook.Cli;

// Outlook COM interop requires an STA thread.
int exitCode = 0;
var sta = new Thread(() =>
{
    var rootCommand = new RootCommand("outlook — Outlook CLI for email, calendar, contacts and calendar sync");
    rootCommand.Subcommands.Add(AccountsCommand.Build());
    rootCommand.Subcommands.Add(SyncCommand.Build());
    rootCommand.Subcommands.Add(EmailCommand.Build());
    rootCommand.Subcommands.Add(CalendarCommand.Build());
    rootCommand.Subcommands.Add(ContactsCommand.Build());

    exitCode = rootCommand.Parse(args).Invoke();
});
sta.SetApartmentState(ApartmentState.STA);
sta.Start();
sta.Join();
return exitCode;

