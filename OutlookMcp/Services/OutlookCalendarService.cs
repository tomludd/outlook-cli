using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Text.Json;

namespace OutlookMcp.Services;

[SupportedOSPlatform("windows")]
public class OutlookCalendarService : IDisposable
{
    // Outlook constants
    private const int OlFolderCalendar = 9;
    private const int OlAppointmentItem = 1;
    private const int OlMeeting = 1;
    private const int OlBusy = 2;
    private const int OlTentative = 1;
    private const int OlFree = 0;
    private const int OlOutOfOffice = 3;
    private const int OlResponseAccepted = 3;
    private const int OlResponseDeclined = 4;
    private const int OlResponseTentative = 2;
    private const int OlResponseNotResponded = 5;

    private dynamic? _outlookApp;

    private dynamic GetOutlookApp()
    {
        if (_outlookApp != null) return _outlookApp;

        var type = Type.GetTypeFromProgID("Outlook.Application")
            ?? throw new InvalidOperationException(
                "Microsoft Outlook is not installed or not registered on this system.");

        _outlookApp = Activator.CreateInstance(type)
            ?? throw new InvalidOperationException("Failed to create Outlook.Application instance.");

        return _outlookApp;
    }

    private dynamic GetNamespace()
    {
        return GetOutlookApp().GetNamespace("MAPI");
    }

    private dynamic GetStoreFolder(string? account, int folderType)
    {
        var ns = GetNamespace();

        if (string.IsNullOrEmpty(account))
            return ns.GetDefaultFolder(folderType);

        var stores = ns.Stores;
        for (int i = 1; i <= stores.Count; i++)
        {
            var store = stores.Item(i);
            if (string.Equals((string)store.DisplayName, account, StringComparison.OrdinalIgnoreCase))
                return store.GetDefaultFolder(folderType);
        }

        throw new InvalidOperationException($"Account not found: {account}. Use list_accounts to see available accounts.");
    }

    private dynamic GetCalendarFolder(string? account)
    {
        return GetStoreFolder(account, OlFolderCalendar);
    }

    public List<Dictionary<string, object?>> ListEvents(DateTime startDate, DateTime endDate, string? account)
    {
        var calendar = GetCalendarFolder(account);
        var filter = $"[Start] >= '{startDate:M/d/yyyy} 12:00 AM' AND [End] <= '{endDate.AddDays(1):M/d/yyyy} 12:00 AM'";
        var items = calendar.Items.Restrict(filter);
        items.Sort("[Start]");

        var events = new List<Dictionary<string, object?>>();
        for (int i = 1; i <= items.Count; i++)
        {
            var item = items.Item(i);
            events.Add(AppointmentToDict(item));
        }
        return events;
    }

    public string CreateEvent(string subject, DateTime startDateTime, DateTime endDateTime,
        string? location, string? body, bool isMeeting, string? attendees, string? account)
    {
        var calendar = GetCalendarFolder(account);
        var appointment = calendar.Items.Add(OlAppointmentItem);

        appointment.Subject = subject;
        appointment.Start = startDateTime;
        appointment.End = endDateTime;

        if (!string.IsNullOrEmpty(location))
            appointment.Location = location;
        if (!string.IsNullOrEmpty(body))
            appointment.Body = body;

        if (isMeeting && !string.IsNullOrEmpty(attendees))
        {
            appointment.MeetingStatus = OlMeeting;
            foreach (var email in attendees.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
            {
                var recipient = appointment.Recipients.Add(email);
                recipient.Type = 1; // Required attendee
            }
            appointment.Send();
        }
        else
        {
            appointment.Save();
        }

        string entryId = appointment.EntryID;
        Marshal.ReleaseComObject(appointment);
        return entryId;
    }

    public bool UpdateEvent(string eventId, string? subject, DateTime? startDateTime, DateTime? endDateTime,
        string? location, string? body, string? account)
    {
        var ns = GetNamespace();
        dynamic appointment;
        try
        {
            appointment = ns.GetItemFromID(eventId);
        }
        catch
        {
            throw new InvalidOperationException($"Event not found with ID: {eventId}");
        }

        if (appointment == null)
            throw new InvalidOperationException($"Event not found with ID: {eventId}");

        if (!string.IsNullOrEmpty(subject))
            appointment.Subject = subject;
        if (startDateTime.HasValue)
            appointment.Start = startDateTime.Value;
        if (endDateTime.HasValue)
            appointment.End = endDateTime.Value;
        if (!string.IsNullOrEmpty(location))
            appointment.Location = location;
        if (!string.IsNullOrEmpty(body))
            appointment.Body = body;

        appointment.Save();
        Marshal.ReleaseComObject(appointment);
        return true;
    }

    public bool DeleteEvent(string eventId, string? account)
    {
        var ns = GetNamespace();
        dynamic appointment;
        try
        {
            appointment = ns.GetItemFromID(eventId);
        }
        catch
        {
            throw new InvalidOperationException($"Event not found with ID: {eventId}");
        }

        if (appointment == null)
            throw new InvalidOperationException($"Event not found with ID: {eventId}");

        appointment.Delete();
        Marshal.ReleaseComObject(appointment);
        return true;
    }

    public List<Dictionary<string, string>> FindFreeSlots(DateTime startDate, DateTime endDate,
        int durationMinutes = 30, int workDayStart = 9, int workDayEnd = 17, string? account = null)
    {
        // Get all events in range
        var calendar = GetCalendarFolder(account);
        var filter = $"[Start] >= '{startDate:M/d/yyyy} 12:00 AM' AND [End] <= '{endDate.AddDays(1):M/d/yyyy} 12:00 AM'";
        var items = calendar.Items.Restrict(filter);
        items.Sort("[Start]");

        // Collect busy slots (only Busy or OutOfOffice)
        var busySlots = new List<(DateTime Start, DateTime End)>();
        for (int i = 1; i <= items.Count; i++)
        {
            var item = items.Item(i);
            int busyStatus = (int)item.BusyStatus;
            if (busyStatus == OlBusy || busyStatus == OlOutOfOffice)
            {
                busySlots.Add(((DateTime)item.Start, (DateTime)item.End));
            }
        }

        // Find free slots
        var freeSlots = new List<Dictionary<string, string>>();
        var currentDate = startDate.Date;
        while (currentDate <= endDate.Date)
        {
            // Skip weekends
            if (currentDate.DayOfWeek != DayOfWeek.Saturday && currentDate.DayOfWeek != DayOfWeek.Sunday)
            {
                var slotStart = currentDate.AddHours(workDayStart);
                var dayEnd = currentDate.AddHours(workDayEnd);

                while (slotStart.AddMinutes(durationMinutes) <= dayEnd)
                {
                    var slotEnd = slotStart.AddMinutes(durationMinutes);
                    bool isFree = !busySlots.Any(b => slotStart < b.End && slotEnd > b.Start);

                    if (isFree)
                    {
                        freeSlots.Add(new Dictionary<string, string>
                        {
                            ["start"] = slotStart.ToString("M/d/yyyy h:mm tt"),
                            ["end"] = slotEnd.ToString("M/d/yyyy h:mm tt")
                        });
                    }

                    slotStart = slotStart.AddMinutes(30); // 30-minute increments
                }
            }
            currentDate = currentDate.AddDays(1);
        }

        return freeSlots;
    }

    public Dictionary<string, object?> GetAttendeeStatus(string eventId, string? account)
    {
        var ns = GetNamespace();
        dynamic appointment;
        try
        {
            appointment = ns.GetItemFromID(eventId);
        }
        catch
        {
            throw new InvalidOperationException($"Event not found with ID: {eventId}");
        }

        if ((int)appointment.MeetingStatus != OlMeeting)
            throw new InvalidOperationException("The specified event is not a meeting.");

        var attendees = new List<Dictionary<string, string>>();
        var recipients = appointment.Recipients;
        for (int i = 1; i <= recipients.Count; i++)
        {
            var recipient = recipients.Item(i);
            var responseStatus = (int)recipient.MeetingResponseStatus switch
            {
                OlResponseAccepted => "Accepted",
                OlResponseDeclined => "Declined",
                OlResponseTentative => "Tentative",
                OlResponseNotResponded => "Not Responded",
                _ => "Unknown"
            };

            attendees.Add(new Dictionary<string, string>
            {
                ["name"] = (string)recipient.Name,
                ["email"] = (string)recipient.Address,
                ["responseStatus"] = responseStatus
            });
        }

        var result = new Dictionary<string, object?>
        {
            ["subject"] = (string)appointment.Subject,
            ["start"] = ((DateTime)appointment.Start).ToString("M/d/yyyy h:mm tt"),
            ["end"] = ((DateTime)appointment.End).ToString("M/d/yyyy h:mm tt"),
            ["location"] = (string)appointment.Location,
            ["organizer"] = (string)appointment.Organizer,
            ["attendees"] = attendees
        };

        Marshal.ReleaseComObject(appointment);
        return result;
    }

    public List<Dictionary<string, object>> GetCalendars()
    {
        var ns = GetNamespace();
        var calendars = new List<Dictionary<string, object>>(); 

        var stores = ns.Stores;
        for (int i = 1; i <= stores.Count; i++)
        {
            var store = stores.Item(i);
            try
            {
                var calendarFolder = store.GetDefaultFolder(OlFolderCalendar);
                if (calendarFolder != null)
                {
                    calendars.Add(new Dictionary<string, object>
                    {
                        ["name"] = (string)store.DisplayName,
                        ["isDefault"] = i == 1
                    });
                }
            }
            catch
            {
                // No calendar folder in this store — skip
            }
        }

        return calendars;
    }

    public List<Dictionary<string, object>> ListAccounts()
    {
        var ns = GetNamespace();
        var accounts = new List<Dictionary<string, object>>();

        var stores = ns.Stores;
        for (int i = 1; i <= stores.Count; i++)
        {
            var store = stores.Item(i);
            accounts.Add(new Dictionary<string, object>
            {
                ["displayName"] = (string)store.DisplayName,
                ["storeId"] = (string)store.StoreID,
                ["isDefault"] = i == 1
            });
        }

        return accounts;
    }

    private Dictionary<string, object?> AppointmentToDict(dynamic appointment)
    {
        var dict = new Dictionary<string, object?>
        {
            ["id"] = (string)appointment.EntryID,
            ["subject"] = (string)appointment.Subject,
            ["start"] = ((DateTime)appointment.Start).ToString("M/d/yyyy h:mm tt"),
            ["end"] = ((DateTime)appointment.End).ToString("M/d/yyyy h:mm tt"),
            ["location"] = (string)appointment.Location,
            ["body"] = (string)appointment.Body,
            ["organizer"] = (string)appointment.Organizer,
            ["isRecurring"] = (bool)appointment.IsRecurring,
            ["isMeeting"] = (int)appointment.MeetingStatus == OlMeeting
        };

        // Busy status
        dict["busyStatus"] = (int)appointment.BusyStatus switch
        {
            OlBusy => "Busy",
            OlTentative => "Tentative",
            OlFree => "Free",
            OlOutOfOffice => "Out of Office",
            _ => "Unknown"
        };

        // Attendees
        var attendees = new List<Dictionary<string, string>>();
        if ((int)appointment.MeetingStatus == OlMeeting)
        {
            var recipients = appointment.Recipients;
            for (int i = 1; i <= recipients.Count; i++)
            {
                var recipient = recipients.Item(i);
                var responseStatus = (int)recipient.MeetingResponseStatus switch
                {
                    OlResponseAccepted => "Accepted",
                    OlResponseDeclined => "Declined",
                    OlResponseTentative => "Tentative",
                    OlResponseNotResponded => "Not Responded",
                    _ => "Unknown"
                };

                attendees.Add(new Dictionary<string, string>
                {
                    ["name"] = (string)recipient.Name,
                    ["email"] = (string)recipient.Address,
                    ["responseStatus"] = responseStatus
                });
            }
        }
        dict["attendees"] = attendees;

        return dict;
    }

    public void Dispose()
    {
        if (_outlookApp != null)
        {
            Marshal.ReleaseComObject(_outlookApp);
            _outlookApp = null;
        }
    }
}
