# Outlook.MCP

Global dotnet tool that exposes Microsoft Outlook as an MCP (Model Context Protocol) server over stdio. Allows AI assistants such as GitHub Copilot to read and manage mail, calendar, and contacts in your local Outlook.

## Install

Pack and install from source:

```powershell
dotnet pack Outlook.MCP/Outlook.MCP.csproj -c Release -o nupkg
dotnet tool install --global outlook-mcp --add-source ./nupkg
```

Update an existing installation:

```powershell
dotnet pack Outlook.MCP/Outlook.MCP.csproj -c Release -o nupkg
dotnet tool update --global outlook-mcp --add-source ./nupkg
```

Verify:

```powershell
outlook-mcp --version
```

## MCP configuration

Add to your `mcp.json`:

```json
{
  "servers": {
    "outlook-mcp": {
      "type": "stdio",
      "command": "outlook-mcp",
      "args": []
    }
  }
}
```

## Tools

### Account

| Tool | Description |
|------|-------------|
| `list_accounts` | List all available Outlook accounts/stores |

### Calendar

| Tool | Description |
|------|-------------|
| `list_events` | List events in a date range |
| `get_event` | Get full event details including body and attendees |
| `create_event` | Create a new event or meeting |
| `update_event` | Update an existing event |
| `delete_event` | Delete an event by ID |
| `find_free_slots` | Find available time slots |
| `get_attendee_status` | Check attendee response status |
| `get_calendars` | List available calendars |

### Mail

| Tool | Description |
|------|-------------|
| `list_emails` | List emails from a folder with optional filters |
| `get_email` | Get full email including body |
| `search_emails` | Search by keyword across subject, body, and sender |
| `send_email` | Send a new email |
| `reply_to_email` | Reply to an existing email |
| `forward_email` | Forward an email to new recipients |

### Contacts

| Tool | Description |
|------|-------------|
| `list_contacts` | List contacts |
| `search_contacts` | Search by name, email, or company |
| `get_contact` | Get full contact details |
| `create_contact` | Create a new contact |
| `update_contact` | Update an existing contact |
| `delete_contact` | Delete a contact |

## Date/Time format

- Dates: `yyyy-MM-dd` (e.g. `2026-03-30`)
- Times: `HH:mm` 24-hour (e.g. `14:00`)

## Dependencies

- [Outlook.COM](../Outlook.COM/readme.md) — COM interop services
- `ModelContextProtocol` 1.1.0
- `Microsoft.Extensions.Hosting` 9.0.6
