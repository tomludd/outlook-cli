# outlook-cli

A global .NET CLI tool for Microsoft Outlook on Windows. Manage email, calendar events, and contacts directly from the terminal, and sync busy time between multiple Outlook accounts.

> Communicates with Outlook via the **COM API** — requires Windows with Microsoft Outlook installed and running. All output is JSON.

## ✨ Features

### 📧 Email

```powershell
outlook email list [--folder inbox|sent|drafts|outbox] [--count 20] [--subject <filter>] [--sender <email>] [--account <name>] [--after yyyy-MM-dd] [--before yyyy-MM-dd] [--body]
outlook email get <id>
outlook email search <query> [--max 20] [--account <name>] [--body]
outlook email send --to <email> --subject <text> --body <text> [--cc <email>] [--bcc <email>] [--html] [--importance low|normal|high] [--attach <path>] [--account <name>]
outlook email reply <id> --body <text> [--all]
outlook email forward <id> --to <email> [--body <text>]
```

### 📅 Calendar

```powershell
outlook calendar list <yyyy-MM-dd> <yyyy-MM-dd> [--account <name>] [--include-blocked]
outlook calendar get <id>
outlook calendar create --subject <text> --start-date yyyy-MM-dd --start-time HH:mm [--end-date yyyy-MM-dd] [--end-time HH:mm] [--location <text>] [--body <text>] [--meeting] [--attendees <email;email>] [--account <name>]
outlook calendar update <id> [--subject <text>] [--start-date yyyy-MM-dd] [--start-time HH:mm] [--end-date yyyy-MM-dd] [--end-time HH:mm] [--location <text>] [--body <text>] [--account <name>]
outlook calendar delete <id> [--account <name>]
outlook calendar free-slots <yyyy-MM-dd> [--to yyyy-MM-dd] [--duration 30] [--work-start 9] [--work-end 17] [--account <name>]
outlook calendar attendees <id> [--account <name>]
outlook calendar calendars
outlook calendar respond <id> <accept|decline|tentative> [--account <name>]
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
outlook sync --source <account> --target <account> [--from yyyy-MM-dd] [--to yyyy-MM-dd] [--mode block|copy] [--outside-hours] [--busy-name <text>] [--busy-location <text>]
outlook sync purge --account <account> [--from yyyy-MM-dd] [--to yyyy-MM-dd]
```

| Option | Default | Description |
|--------|---------|-------------|
| `--source` | required | Account to sync events **from** |
| `--target` | required | Account to sync events **to** |
| `--from` | today | Start date |
| `--to` | today + 90 days | End date |
| `--mode` | `block` | `block` — anonymous Busy/OOO placeholders · `copy` — copies title and description (shows as Free) |
| `--outside-hours` | false | Only sync events outside 07:00–18:00 |
| `--busy-name` | `Busy` | Subject used for `block` mode placeholders (OOO events keep "Out of Office") |
| `--busy-location` | none | Location used for `block` mode placeholders. Omit for no location. `copy` mode is unaffected. |

The Reminder App's sync rule editor exposes both fields per rule.

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

## ⚙️ Outlook configuration (Programmatic Access warning)

When `outlook-cli` talks to Outlook through the COM API, Outlook may pop up a security prompt:

> **A program is trying to access e-mail address information stored in Outlook.**

This blocks the command until you click **Allow**, which is a problem for unattended use (e.g. the reminder app or scheduled sync). To suppress the prompt, tell Outlook to never warn about programmatic access.

### Never warn me about suspicious activity

1. Close Outlook completely.
2. Start Outlook **as administrator** (right-click the Outlook shortcut → **Run as administrator**).
   > The *Programmatic Access* radio buttons are greyed out unless Outlook is running elevated, so this step is required.
3. Go to **File → Options → Trust Center**.
4. Click **Trust Center Settings…** on the right.
5. Select **Programmatic Access** in the left list.
6. Under *Outlook Security settings*, choose:
   **Never warn me about suspicious activity (not recommended)**.
7. Click **OK** twice, then restart Outlook normally.

> **Note:** This setting is stored per-Outlook-profile and is the reason the option is only editable while running as administrator. On some setups (e.g. Windows ARM / Parallels) Outlook reports *Antivirus status: Invalid*, which is exactly why the option above is needed.

Reference: [A program is trying to access email address information stored in Outlook](https://learn.microsoft.com/en-us/answers/questions/4540111/a-program-is-trying-to-access-email-address-inform)

---