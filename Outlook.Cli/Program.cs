using System.CommandLine;
using Outlook.Cli;

var rootCommand = new RootCommand("outlook — Outlook CLI for email, calendar, contacts and calendar sync");
rootCommand.Subcommands.Add(AccountsCommand.Build());
rootCommand.Subcommands.Add(SyncCommand.Build());
rootCommand.Subcommands.Add(EmailCommand.Build());
rootCommand.Subcommands.Add(CalendarCommand.Build());
rootCommand.Subcommands.Add(ContactsCommand.Build());

return rootCommand.Parse(args).Invoke();


