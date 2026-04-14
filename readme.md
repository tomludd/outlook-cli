# outlook-cli

A global .NET CLI tool for Microsoft Outlook on Windows. Manage email, calendar events, and contacts directly from the terminal, and sync busy time between multiple Outlook accounts.

> Communicates with Outlook via the **COM API** — requires Windows with Microsoft Outlook installed and running. All output is JSON.

## ✨ Features

### 📧 Email

```powershell
outlook email list [--folder inbox|sent|drafts|outbox] [--count 20] [--subject <filter>] [--sender <email>] [--account <name>] [--after yyyy-MM-dd] [--before yyyy-MM-dd]
outlook email get <id>
outlook email search <query> [--max 20] [--account <name>]
outlook email send --to <email> --subject <text> --body <text> [--cc <email>] [--bcc <email>] [--html] [--importance low|normal|high] [--attach <path>] [--account <name>]
outlook email reply <id> --body <text> [--all]
outlook email forward <id> --to <email> [--body <text>]
```

### 📅 Calendar

```powershell
outlook calendar list <yyyy-MM-dd> <yyyy-MM-dd> [--account <name>]
outlook calendar get <id>
outlook calendar create --subject <text> --start-date yyyy-MM-dd --start-time HH:mm [--end-date yyyy-MM-dd] [--end-time HH:mm] [--location <text>] [--body <text>] [--meeting] [--attendees <email;email>] [--account <name>]
outlook calendar update <id> [--subject <text>] [--start-date yyyy-MM-dd] [--start-time HH:mm] [--end-date yyyy-MM-dd] [--end-time HH:mm] [--location <text>] [--body <text>] [--account <name>]
outlook calendar delete <id> [--account <name>]
outlook calendar free-slots <yyyy-MM-dd> [--to yyyy-MM-dd] [--duration 30] [--work-start 9] [--work-end 17] [--account <name>]
outlook calendar attendees <id> [--account <name>]
outlook calendar calendars
```

### 👤 Contacts

```powershell
outlook contacts list [--count 50] [--account <name>]
outlook contacts search <query> [--max 20] [--account <name>]
outlook contacts get <id>
outlook contacts create [--first <name>] [--last <name>] [--email <email>] [--phone <number>] [--mobile <number>] [--company <name>] [--title <title>] [--address <text>] [--notes <text>] [--account <name>]
outlook contacts update <id> [--first <name>] [--last <name>] [--email <email>] [--phone <number>] [--mobile <number>] [--company <name>] [--title <title>] [--address <text>] [--notes <text>]
outlook contacts delete <id>
```

### 🔄 Calendar sync

Syncs busy time between Outlook calendars. Blocking events are tagged with a hidden marker and never re-synced, preventing cascading blocks.

```powershell
outlook sync --source <account> --target <account> [--from yyyy-MM-dd] [--to yyyy-MM-dd] [--mode block|copy] [--outside-hours]
```

| Option | Default | Description |
|--------|---------|-------------|
| `--source` | required | Account to sync events **from** |
| `--target` | required | Account to sync events **to** |
| `--from` | today | Start date |
| `--to` | today + 90 days | End date |
| `--mode` | `block` | `block` — anonymous Busy/OOO placeholders · `copy` — copies title and description (shows as Free) |
| `--outside-hours` | false | Only sync events outside 07:00–18:00 |

```powershell
# Block busy time both ways between two work accounts
outlook sync --source "work@company.com" --target "me@personal.com"
outlook sync --source "me@personal.com" --target "work@company.com"

# Copy outside-hours events (with title + description) to personal calendar
outlook sync --source "work@company.com" --target "me@personal.com" --mode copy --outside-hours
```

### 🏦 Accounts

List account display names for use with `--account`:

```powershell
outlook accounts
```

---

## 📦 Install

### From NuGet.org

```powershell
dotnet tool install --global outlook-cli
```

Or run without installing via `dnx`:

```powershell
dnx outlook-cli
```

### From source

```powershell
dotnet pack Outlook.Cli/Outlook.Cli.csproj -c Release -o nupkg
dotnet tool install --global outlook-cli --add-source ./nupkg
```

### 🗑️ Uninstall

```powershell
dotnet tool uninstall --global outlook-cli
```

---