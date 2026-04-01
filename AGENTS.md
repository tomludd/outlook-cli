# outlook-mcp — Agent Instructions

This is a .NET 10 Model Context Protocol (MCP) server for Microsoft Outlook on Windows, distributed as a global dotnet tool (`outlook-mcp`).

## Repo layout

| Path | Purpose |
|------|---------|
| `OutlookMcp/` | Main server project |
| `OutlookMcp/Services/` | Outlook COM interop service wrappers |
| `OutlookMcp/Tools/` | MCP tool definitions (one file per capability area) |
| `OutlookMcp.IntegrationTests/` | Integration test project |
| `nupkg/` | Output folder for packed NuGet packages |

## Build & run

```powershell
dotnet build
```

## Pack and install (after code changes)

Run these commands from the repo root whenever you change `OutlookMcp`:

```powershell
# Pack a new .nupkg
dotnet pack OutlookMcp/OutlookMcp.csproj -c Release -o nupkg

# First-time install
dotnet tool install --global outlook-mcp --add-source ./nupkg

# Subsequent updates
dotnet tool update --global outlook-mcp --add-source ./nupkg
```

Verify:

```powershell
outlook-mcp --version
```

## MCP server configuration

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

## Project conventions

- Target framework: `net10.0` (platform-specific TFM `net10.0-windows` is intentionally avoided; it is incompatible with `PackAsTool`).
- `NoWarn CA1416` suppresses Windows-platform analyzer warnings that are intentional for an Outlook COM interop tool.
- Tool command name: `outlook-mcp` (set via `<ToolCommandName>` in the `.csproj`).
- Package ID / version: `outlook-mcp` / `1.0.0` — bump `<Version>` in `OutlookMcp.csproj` before re-packing a release.
- Outlook COM interop runs in-process; Outlook must be installed and running on Windows.
