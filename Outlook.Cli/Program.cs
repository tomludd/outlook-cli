using System.CommandLine;
using Outlook.Cli;

var rootCommand = new RootCommand("outlook — Outlook CLI for email, calendar, contacts and calendar sync");
rootCommand.AddCommand(AccountsCommand.Build());
rootCommand.AddCommand(SyncCommand.Build());
rootCommand.AddCommand(EmailCommand.Build());
rootCommand.AddCommand(CalendarCommand.Build());
rootCommand.AddCommand(ContactsCommand.Build());

return await rootCommand.InvokeAsync(args);
