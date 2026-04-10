# outlook-mcp

A collection of .NET 10 libraries and tools for integrating with Microsoft Outlook on Windows via COM interop.

## Projects

| Project | Description |
|---------|-------------|
| [Outlook.COM](Outlook.COM/readme.md) | Class library — Outlook COM interop services (mail, calendar, contacts) |
| [Outlook.MCP](Outlook.MCP/readme.md) | Global dotnet tool — MCP server exposing Outlook as AI tools |
| [Outlook.COM.IntegrationTests](Outlook.COM.IntegrationTests/readme.md) | Integration tests for `Outlook.COM` |
| [Outlook.MCP.IntegrationTests](Outlook.MCP.IntegrationTests/readme.md) | Integration tests for `Outlook.MCP` |

## Prerequisites

- Windows
- Microsoft Outlook desktop client installed and running
- .NET 10.0 SDK

## Build

```powershell
dotnet build
```

## Security

All operations are local — no data is sent to external servers. On first use, Outlook may prompt you to allow programmatic access.

## License

MIT
