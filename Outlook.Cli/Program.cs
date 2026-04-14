using System.CommandLine;
using Outlook.Cli;

var rootCommand = new RootCommand("outlook-sync — Calendar blocking sync tool");
rootCommand.AddCommand(AccountsCommand.Build());
rootCommand.AddCommand(SyncCommand.Build());

return await rootCommand.InvokeAsync(args);
