# outlook-cli — Agent Instructions

This is a .NET 10 CLI tool and Windows reminder app for Microsoft Outlook on Windows, distributed as a global dotnet tool (`outlook-cli`).

## Repo layout

| Path | Purpose |
|------|---------|
| `Outlook.COM/` | Core Outlook COM interop library |
| `Outlook.COM/OutlookCalendarService.cs` | Calendar operations via COM API |
| `Outlook.COM/OutlookContactService.cs` | Contact management via COM API |
| `Outlook.COM/OutlookMailService.cs` | Email operations via COM API |
| `Outlook.Cli/` | CLI tool project (`outlook-cli`) |
| `Outlook.Cli/AccountsCommand.cs` | List Outlook accounts |
| `Outlook.Cli/CalendarCommand.cs` | Calendar CLI commands |
| `Outlook.Cli/ContactsCommand.cs` | Contact CLI commands |
| `Outlook.Cli/EmailCommand.cs` | Email CLI commands |
| `Outlook.Cli/SyncCommand.cs` | Calendar sync between accounts |
| `Outlook.Cli/CalendarSyncService.cs` | Calendar sync logic |
| `Outlook.ReminderApp/` | Windows Forms meeting reminder widget |
| `Outlook.ReminderApp/MainForm.cs` | UI for reminder widget (supports display change events) |
| `Outlook.ReminderApp/MeetingReminderService.cs` | Meeting notification logic |
| `Outlook.ReminderApp/MeetingActionStateStore.cs` | Tracks dismissed reminders |
| `Outlook.ReminderApp/TeamsJoinLinkResolver.cs` | Extracts Teams join URLs |
| `Outlook.COM.IntegrationTests/` | Integration test project |
| `nupkg/` | Output folder for packed NuGet packages |
| `artifacts/` | Build output artifacts |

## Build & run

```powershell
dotnet build
```

## Pack and install outlook-cli (after code changes)

Run these commands from the repo root whenever you change `Outlook.Cli`:

> **Always bump `<Version>` in `Outlook.Cli/Outlook.Cli.csproj` before packing.** `dotnet tool update` requires a higher version number to pick up changes.

```powershell
# Pack a new .nupkg
dotnet pack Outlook.Cli/Outlook.Cli.csproj -c Release -o nupkg

# First-time install
dotnet tool install --global outlook-cli --add-source ./nupkg

# Subsequent updates
dotnet tool update --global outlook-cli --add-source ./nupkg
```

## Run ReminderApp

> **ReminderApp is always running.** Kill the process before building or the build will fail (locked exe).

```powershell
Stop-Process -Name "Outlook.ReminderApp" -Force -ErrorAction SilentlyContinue
```

```powershell
dotnet run --project Outlook.ReminderApp
```

Or build and run the executable directly:

```powershell
Stop-Process -Name "Outlook.ReminderApp" -Force -ErrorAction SilentlyContinue
dotnet build Outlook.ReminderApp -c Release
Start-Process .\Outlook.ReminderApp\bin\Release\net10.0-windows\Outlook.ReminderApp.exe
```

## Project conventions

- Target framework: 
  - `Outlook.COM`, `Outlook.Cli`: `net10.0` (platform-specific TFM `net10.0-windows` is intentionally avoided; it is incompatible with `PackAsTool`)
  - `Outlook.ReminderApp`: `net10.0-windows` (requires WinForms)
- `NoWarn CA1416` suppresses Windows-platform analyzer warnings that are intentional for Outlook COM interop.
- CLI tool command name: `outlook-cli` (set via `<ToolCommandName>` in `Outlook.Cli.csproj`)
- Outlook COM interop runs in-process; Outlook must be installed and running on Windows.
- ReminderApp automatically repositions itself when display configuration changes (monitor plug/unplug, laptop close/open)
