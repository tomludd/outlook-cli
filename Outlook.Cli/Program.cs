using System.CommandLine;
using Outlook.Cli;

// Outlook COM work runs on Outlook.COM's own dedicated STA worker thread (see ComTimeout);
// this entry point doesn't touch COM objects directly, so it doesn't need to be STA itself.
var rootCommand = new RootCommand("outlook — Outlook CLI for email, calendar, contacts and calendar sync");
rootCommand.Subcommands.Add(AccountsCommand.Build());
rootCommand.Subcommands.Add(SyncCommand.Build());
rootCommand.Subcommands.Add(EmailCommand.Build());
rootCommand.Subcommands.Add(CalendarCommand.Build());
rootCommand.Subcommands.Add(ContactsCommand.Build());

return rootCommand.Parse(args).Invoke();

