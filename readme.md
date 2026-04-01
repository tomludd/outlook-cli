# Outlook Calendar MCP Server (C#)

A Model Context Protocol (MCP) server that allows AI assistants to access and manage your local Microsoft Outlook calendar. Windows only.

This is a C# reimplementation of [merajmehrabi/Outlook_Calendar_MCP](https://github.com/merajmehrabi/Outlook_Calendar_MCP). Instead of VBScript, it uses direct COM interop — no VBScript dependency, no deprecation concerns.

## Prerequisites

- Windows
- Microsoft Outlook desktop client installed
- .NET 10.0 SDK

## Install as a global dotnet tool

Pack and install (or update) the tool from source:

```powershell
dotnet pack OutlookMcp/OutlookMcp.csproj -c Release -o nupkg
dotnet tool install --global outlook-mcp --add-source ./nupkg
```

If you already have a previous version installed, update instead:

```powershell
dotnet pack OutlookMcp/OutlookMcp.csproj -c Release -o nupkg
dotnet tool update --global outlook-mcp --add-source ./nupkg
```

Verify the installation:

```powershell
outlook-mcp --version
```

## MCP Configuration

### VS Code (GitHub Copilot / Cline)

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
