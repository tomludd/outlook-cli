# Outlook.COM.IntegrationTests

Integration tests for [Outlook.COM](../Outlook.COM/readme.md). Tests run against a live Outlook instance on Windows.

## Prerequisites

- Windows
- Microsoft Outlook desktop client installed and running with at least one account configured
- .NET 10.0 SDK

## Run

```powershell
dotnet test Outlook.COM.IntegrationTests/Outlook.COM.IntegrationTests.csproj
```

## Test classes

| Class | Covers |
|-------|--------|
| `AccountTests` | `OutlookCalendarService.ListAccounts` |
| `CalendarTests` | `OutlookCalendarService` — list, find free slots, date range filtering |
| `ContactTests` | `OutlookContactService` — list, search, get |
| `EmailTests` | `OutlookMailService` — list, get, search, filter |

All test classes share a single `OutlookFixture` instance via xunit's `IClassFixture<T>` to avoid redundant COM initialization.
