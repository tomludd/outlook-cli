# Outlook Calendar MCP Server (C#)

A Model Context Protocol (MCP) server that allows AI assistants to access and manage your local Microsoft Outlook calendar. Windows only.

This is a C# reimplementation of [merajmehrabi/Outlook_Calendar_MCP](https://github.com/merajmehrabi/Outlook_Calendar_MCP). Instead of VBScript, it uses direct COM interop — no VBScript dependency, no deprecation concerns.

## Prerequisites

- Windows
- Microsoft Outlook desktop client installed
- .NET 9.0 SDK

## Build

```bash
dotnet build
```

## MCP Configuration

### VS Code (GitHub Copilot / Cline)

Add to your MCP settings:

```json
{
  "servers": {
    "outlook-calendar": {
      "type": "stdio",
      "command": "dotnet",
      "args": ["run", "--project", "D:\\outlook-mcp\\OutlookMcp"]
    }
  }
}
```

Or if you publish a self-contained exe:

```json
{
  "servers": {
    "outlook-calendar": {
      "type": "stdio",
      "command": "D:\\outlook-mcp\\OutlookMcp\\bin\\Release\\net9.0\\OutlookMcp.exe"
    }
  }
}
```

### Claude Desktop

Add to `%APPDATA%\Claude\claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "outlook-calendar": {
      "command": "dotnet",
      "args": ["run", "--project", "D:\\outlook-mcp\\OutlookMcp"]
    }
  }
}
```

## Tools

| Tool | Description |
|------|-------------|
| `list_events` | List calendar events within a date range |
| `create_event` | Create a new event or meeting |
| `update_event` | Update an existing event |
| `delete_event` | Delete an event by ID |
| `find_free_slots` | Find available time slots |
| `get_attendee_status` | Check attendee response status |
| `get_calendars` | List available calendars |

### Date/Time Format

- Dates: `yyyy-MM-dd` (e.g. `2026-03-30`)
- Times: `HH:mm` 24-hour format (e.g. `14:00`)

## Security

- All operations are local — no data is sent to external servers.
- On first use, Outlook may prompt you to allow programmatic access.

## License

MIT
