# Outlook.MCP.IntegrationTests

Integration tests for [Outlook.MCP](../Outlook.MCP/readme.md) tool-layer classes. Tests run against a live Outlook instance on Windows.

## Prerequisites

- Windows
- Microsoft Outlook desktop client installed and running with at least one account configured
- .NET 10.0 SDK

## Run

```powershell
dotnet test Outlook.MCP.IntegrationTests/Outlook.MCP.IntegrationTests.csproj
```
