# Outlook.COM

.NET 10 class library providing Outlook COM interop services for mail, calendar, and contacts. Windows only.

Add a project reference to use it in your own project:

```xml
<ProjectReference Include="..\Outlook.COM\Outlook.COM.csproj" />
```

## Namespace

`Outlook.COM`

## Services

### `OutlookCalendarService`

| Method | Description |
|--------|-------------|
| `ListEvents(start, end, account?)` | List events in a date range |
| `GetEvent(eventId)` | Get full event details including body and attendees |
| `CreateEvent(subject, start, end, ...)` | Create an event or meeting |
| `UpdateEvent(eventId, ...)` | Update an existing event |
| `DeleteEvent(eventId, account?)` | Delete an event |
| `FindFreeSlots(start, end, duration, ...)` | Find available time slots |
| `GetAttendeeStatus(eventId, account?)` | Get per-attendee response status |
| `GetCalendars()` | List available calendars |
| `ListAccounts()` | List available Outlook stores/accounts |

### `OutlookMailService`

| Method | Description |
|--------|-------------|
| `ListEmails(folder?, count, ...)` | List emails with optional filters |
| `GetEmail(entryId)` | Get full email including body |
| `SendEmail(to, subject, body, ...)` | Send a new email |
| `ReplyToEmail(entryId, body, replyAll)` | Reply to an email |
| `ForwardEmail(entryId, to, body?)` | Forward an email |
| `SearchEmails(query, maxResults, account?)` | Search by subject, body, or sender |

### `OutlookContactService`

| Method | Description |
|--------|-------------|
| `ListContacts(count, account?)` | List contacts sorted by name |
| `SearchContacts(query, maxResults, account?)` | Search by name, email, or company |
| `GetContact(entryId)` | Get full contact details |
| `CreateContact(firstName, lastName, ...)` | Create a new contact |
| `UpdateContact(entryId, ...)` | Update an existing contact |
| `DeleteContact(entryId)` | Delete a contact |

## Prerequisites

- Windows
- Microsoft Outlook desktop client installed and running
- .NET 10.0 SDK

## Notes

- Each service manages its own Outlook COM handle. Instantiate with `new` and dispose when done (`using var svc = new OutlookCalendarService()`).
- The `account` parameter on most methods accepts an Outlook store display name (e.g. `tommy@example.com`). Omit to operate across all accounts.
- Date strings use `yyyy-MM-dd`, times use `HH:mm` (24-hour).
