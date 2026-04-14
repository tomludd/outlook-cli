using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Globalization;
using System.Text.Json;

namespace Outlook.COM;

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

    public List<Dictionary<string, object?>> ListEvents(DateTime startDate, DateTime endDate, string? account, int bodyLength = 50)
    {
        var events = new List<Dictionary<string, object?>>();

        if (string.IsNullOrEmpty(account))
        {
            var ns = GetNamespace();
            var stores = ns.Stores;
            for (int i = 1; i <= stores.Count; i++)
            {
                try { CollectEvents(stores.Item(i).GetDefaultFolder(OlFolderCalendar), startDate, endDate, events, bodyLength); }
                catch { /* Store has no calendar folder */ }
            }
            events.Sort((a, b) => string.Compare(a["start"]?.ToString(), b["start"]?.ToString(), StringComparison.Ordinal));
        }
        else
        {
            CollectEvents(GetCalendarFolder(account), startDate, endDate, events, bodyLength);
        }

        return events;
    }

    private void CollectEvents(dynamic folder, DateTime startDate, DateTime endDate, List<Dictionary<string, object?>> events, int bodyLength = 50)
    {
        var restrictedItems = GetCalendarItemsInRange(folder, startDate, endDate);
        var item = restrictedItems.GetFirst();
        while (item != null)
        {
            events.Add(AppointmentToDict(item, bodyLength: bodyLength));
            item = restrictedItems.GetNext();
        }
    }

    public string CreateEvent(string subject, DateTime startDateTime, DateTime endDateTime,
        string? location, string? body, bool isMeeting, string? attendees, string? account,
        bool reminderEnabled = true, int busyStatus = OlBusy)
    {
        var calendar = GetCalendarFolder(account);
        var appointment = calendar.Items.Add(OlAppointmentItem);

        appointment.Subject = subject;
        appointment.Start = startDateTime;
        appointment.End = endDateTime;
        appointment.ReminderSet = reminderEnabled;
        appointment.BusyStatus = busyStatus;

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

        // Re-fetch after save: Outlook reassigns EntryID after first save to Exchange
        string tempId = (string)appointment.EntryID;
        Marshal.ReleaseComObject(appointment);
        var ns = GetNamespace();
        dynamic saved = ns.GetItemFromID(tempId);
        string stableId = (string)saved.EntryID;
        Marshal.ReleaseComObject(saved);
        return stableId;
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

        try { appointment.Delete(); }
        catch (System.Runtime.InteropServices.COMException) { /* item deleted but COM reports a move error — ignore */ }
        finally { try { Marshal.ReleaseComObject(appointment); } catch { } }
        return true;
    }

    public List<Dictionary<string, string>> FindFreeSlots(DateTime startDate, DateTime endDate,
        int durationMinutes = 30, int workDayStart = 9, int workDayEnd = 17, string? account = null)
    {
        // Collect busy slots from all relevant calendars
        var busySlots = new List<(DateTime Start, DateTime End)>();

        void CollectBusy(dynamic folder)
        {
            var restrictedItems = GetCalendarItemsInRange(folder, startDate, endDate);
            var item = restrictedItems.GetFirst();
            while (item != null)
            {
                int busyStatus = (int)item.BusyStatus;
                if (busyStatus == OlBusy || busyStatus == OlOutOfOffice)
                    busySlots.Add(((DateTime)item.Start, (DateTime)item.End));
                item = restrictedItems.GetNext();
            }
        }

        if (string.IsNullOrEmpty(account))
        {
            var ns = GetNamespace();
            var stores = ns.Stores;
            for (int i = 1; i <= stores.Count; i++)
            {
                try { CollectBusy(stores.Item(i).GetDefaultFolder(OlFolderCalendar)); }
                catch { /* Store has no calendar folder */ }
            }
        }
        else
        {
            CollectBusy(GetCalendarFolder(account));
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
                            ["start"] = slotStart.ToString("yyyy-MM-dd HH:mm"),
                            ["end"] = slotEnd.ToString("yyyy-MM-dd HH:mm")
                        });
                    }

                    slotStart = slotStart.AddMinutes(30); // 30-minute increments
                }
            }
            currentDate = currentDate.AddDays(1);
        }

        return freeSlots;
    }

    private static dynamic GetCalendarItemsInRange(dynamic folder, DateTime startDate, DateTime endDate)
    {
        var items = folder.Items;
        items.Sort("[Start]");
        items.IncludeRecurrences = true;
        return items.Restrict(BuildDateRangeFilter(startDate, endDate));
    }

    internal static string BuildDateRangeFilter(DateTime startDate, DateTime endDate)
    {
        var rangeStart = startDate.Date;
        var rangeEndExclusive = endDate.Date.AddDays(1);
        return $"[Start] < '{FormatOutlookDateTime(rangeEndExclusive)}' AND [End] > '{FormatOutlookDateTime(rangeStart)}'";
    }

    internal static string FormatOutlookDateTime(DateTime value)
    {
        return value.ToString("g", CultureInfo.CurrentCulture);
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
                ["responseStatus"] = responseStatus
            });
        }

        var result = new Dictionary<string, object?>
        {
            ["subject"] = (string)appointment.Subject,
            ["start"] = ((DateTime)appointment.Start).ToString("yyyy-MM-dd HH:mm"),
            ["end"] = ((DateTime)appointment.End).ToString("yyyy-MM-dd HH:mm"),
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

    private Dictionary<string, object?> AppointmentToDict(dynamic appointment, int bodyLength = 0, bool includeAttendees = false)
    {
        var dict = new Dictionary<string, object?>
        {
            ["id"] = (string)appointment.EntryID,
            ["subject"] = (string)appointment.Subject,
            ["start"] = ((DateTime)appointment.Start).ToString("yyyy-MM-dd HH:mm"),
            ["end"] = ((DateTime)appointment.End).ToString("yyyy-MM-dd HH:mm"),
            ["location"] = (string)appointment.Location,
            ["organizer"] = (string)appointment.Organizer,
            ["isRecurring"] = (bool)appointment.IsRecurring,
            ["isMeeting"] = (int)appointment.MeetingStatus == OlMeeting
        };

        dict["busyStatus"] = (int)appointment.BusyStatus switch
        {
            OlBusy => "Busy",
            OlTentative => "Tentative",
            OlFree => "Free",
            OlOutOfOffice => "Out of Office",
            _ => "Unknown"
        };

        if (bodyLength > 0)
        {
            var fullBody = (string)appointment.Body;
            dict["body"] = fullBody.Length <= bodyLength ? fullBody : fullBody[..bodyLength] + "...";
        }

        if (includeAttendees && (int)appointment.MeetingStatus == OlMeeting)
        {
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
                    ["responseStatus"] = responseStatus
                });
            }
            dict["attendees"] = attendees;
        }

        return dict;
    }

    public Dictionary<string, object?> GetEvent(string eventId)
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

        var result = AppointmentToDict(appointment, bodyLength: int.MaxValue, includeAttendees: true);
        Marshal.ReleaseComObject(appointment);
        return result;
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
