# outlook-sync

A global .NET tool that syncs events between Outlook calendars on Windows. Useful for blocking time across work and personal accounts, or copying outside-hours events to a personal calendar.

## Install

```powershell
dotnet pack Outlook.Cli/Outlook.Cli.csproj -c Release -o nupkg
dotnet tool install --global outlook-sync --add-source ./nupkg
```

To update after code changes (bump `<Version>` in the `.csproj` first):

```powershell
dotnet tool update --global outlook-sync --add-source ./nupkg
```

## Commands

### `accounts`

Lists all Outlook accounts available on this machine.

```powershell
outlook-sync accounts
```

### `sync`

Syncs events from a source calendar to a target calendar.

```powershell
outlook-sync sync --source <account> --target <account> [options]
```

| Option | Required | Description |
|--------|----------|-------------|
| `--source` | Yes | Account to sync events **from** |
| `--target` | Yes | Account to sync events **to** |
| `--from` | No | Start date `yyyy-MM-dd` (default: today) |
| `--to` | No | End date `yyyy-MM-dd` (default: today + 90 days) |
| `--mode` | No | `block` (default) or `copy` — see below |
| `--outside-hours` | No | Only sync events outside 07:00–18:00 |

#### Sync modes

| Mode | Behaviour |
|------|-----------|
| `block` | Creates anonymous **Busy** or **Out of Office** placeholder events in the target calendar |
| `copy` | Copies the original **title and description** to the target (event shows as Free) |

## Examples

Block busy time from work to personal for the next 90 days:

```powershell
outlook-sync sync --source "work@company.com" --target "me@personal.com"
```

Block busy time both ways for a specific day:

```powershell
outlook-sync sync --source "work@company.com" --target "me@personal.com" --from 2026-04-14 --to 2026-04-14
outlook-sync sync --source "me@personal.com" --target "work@company.com" --from 2026-04-14 --to 2026-04-14
```

Copy outside-hours work events (title + description) to personal calendar:

```powershell
outlook-sync sync --source "work@company.com" --target "me@personal.com" --mode copy --outside-hours
```

## How it works

Each synced event is tagged with a hidden marker in its description:

```
[outlook-sync:block:<hash>]
```

The hash is a stable SHA256 of `source:target`, so markers are consistent across every run without any persisted state. On each sync run:

1. Events marked for this source→target pair in the target calendar are discovered.
2. Any marked event whose time slot no longer exists in the source is **deleted**.
3. Any source slot not yet marked in the target is **created**.

`block` and `copy` use separate markers so both modes can run independently on the same account pair.

## Prerequisites

- Windows
- Microsoft Outlook desktop client installed and running
- .NET 10.0 SDK (for building/installing)
