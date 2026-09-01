using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Globalization;
using System.Text.Json;

namespace Outlook.COM;

[SupportedOSPlatform("windows")]
public class OutlookCalendarService
{
    // Outlook constants
    private const int OlFolderCalendar = 9;
    private const int OlAppointmentItem = 1;
    private const int OlMeeting = 1;
    private const int OlMeetingReceived = 3;
    private const int OlMeetingCanceled = 5;
    private const int OlMeetingReceivedAndCanceled = 7;
    private const int OlBusy = 2;
    private const int OlTentative = 1;
    private const int OlFree = 0;
    private const int OlOutOfOffice = 3;
    private const int OlResponseAccepted = 3;
    private const int OlResponseDeclined = 4;
    private const int OlResponseTentative = 2;
    private const int OlResponseNotResponded = 5;

    private dynamic GetOutlookApp() => OutlookComHost.GetApp();

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

    /// <summary>
    /// Returns the SMTP address for a store by matching it to a MAPI account via DeliveryStore.
    /// Falls back to null if no match is found.
    /// </summary>
    private string? GetSmtpAddressForStore(dynamic ns, dynamic store)
    {
        // If the store's DisplayName is already an SMTP address, use it directly.
        try
        {
            var displayName = (string)store.DisplayName;
            if (!string.IsNullOrEmpty(displayName) && displayName.Contains('@'))
                return displayName;
        }
        catch { }

        // For the primary Exchange mailbox (ExchangeStoreType == 1), resolve via CurrentUser.
        try
        {
            if ((int)store.ExchangeStoreType == 1)
            {
                var exchUser = ns.CurrentUser.AddressEntry.GetExchangeUser();
                if (exchUser != null)
                {
                    var smtp = (string)exchUser.PrimarySmtpAddress;
                    if (!string.IsNullOrEmpty(smtp)) return smtp;
                }
            }
        }
        catch { }

        // Otherwise try matching via ns.Accounts DeliveryStore.
        try
        {
            var storeId = (string)store.StoreID;
            var accounts = ns.Accounts;
            for (int i = 1; i <= (int)accounts.Count; i++)
            {
                try
                {
                    var acct = accounts.Item(i);
                    dynamic? deliveryStore = null;
                    try { deliveryStore = acct.DeliveryStore; } catch { }
                    if (deliveryStore is null) continue;
                    if (string.Equals((string)deliveryStore.StoreID, storeId, StringComparison.OrdinalIgnoreCase))
                        return (string)acct.SmtpAddress;
                }
                catch { }
            }
        }
        catch { }
        return null;
    }

    public List<Dictionary<string, object?>> ListEvents(DateTime startDate, DateTime endDate, string? account)
        => OutlookComInvoker.Run(() => ListEventsCore(startDate, endDate, account));

    /// <summary>
    /// Drops the shared cached <c>Outlook.Application</c> RCW so the next COM call reconnects
    /// from scratch. Long-lived callers (e.g. a background poller) that repeatedly restrict a
    /// recurring-items collection with <c>IncludeRecurrences</c> over the same session can hit an
    /// Outlook/MAPI quirk where recurrence expansion silently degrades to a single occurrence;
    /// periodically reconnecting works around it the same way a short-lived process naturally does.
    /// </summary>
    public void ResetConnection() => OutlookComHost.Reset();

    /// <summary>
    /// True for the same busy/stale-connection HRESULTs that <see cref="OutlookComInvoker"/> retries on.
    /// Distinguishes "Outlook rejected this call" from "this store genuinely has no calendar folder".
    /// </summary>
    private static bool IsTransientComFailure(COMException ex) =>
        ex.HResult == RpcHResult.ServerUnavailable
        || ex.HResult == RpcHResult.ServerDisconnected
        || ex.HResult == RpcHResult.CallCanceled
        || ex.HResult == RpcHResult.CallRejected;

    private List<Dictionary<string, object?>> ListEventsCore(DateTime startDate, DateTime endDate, string? account)
    {
        var events = new List<Dictionary<string, object?>>();

        if (string.IsNullOrEmpty(account))
        {
            var ns = GetNamespace();
            var stores = ns.Stores;
            try
            {
                for (int i = 1; i <= stores.Count; i++)
                {
                    dynamic? store = null;
                    try
                    {
                        store = stores.Item(i);
                        var smtpAddress = GetSmtpAddressForStore(ns, store);
                        CollectEvents(store.GetDefaultFolder(OlFolderCalendar), startDate, endDate, events, smtpAddress);
                    }
                    catch (COMException ex) when (IsTransientComFailure(ex))
                    {
                        // Outlook was busy/stale for this one store's call — let it bubble up so
                        // OutlookComInvoker can reset/retry the whole fetch, instead of silently
                        // returning a partial result that's missing this store's events.
                        throw;
                    }
                    catch { /* Store has no calendar folder */ }
                    finally { OutlookComHost.Release(store); }
                }
                events.Sort((a, b) => string.Compare(a["start"]?.ToString(), b["start"]?.ToString(), StringComparison.Ordinal));
            }
            finally
            {
                OutlookComHost.Release(stores);
                OutlookComHost.Release(ns);
            }
        }
        else
        {
            CollectEvents(GetCalendarFolder(account), startDate, endDate, events);
        }

        return events;
    }

    private void CollectEvents(dynamic folder, DateTime startDate, DateTime endDate, List<Dictionary<string, object?>> events, string? accountName = null)
    {
        var restrictedItems = GetCalendarItemsInRange(folder, startDate, endDate);
        try
        {
            dynamic? item = restrictedItems.GetFirst();
            while (item != null)
            {
                var dict = AppointmentToDict(item);
                if (accountName is not null) dict["account"] = accountName;
                events.Add(dict);
                OutlookComHost.Release(item);
                item = restrictedItems.GetNext();
            }
        }
        finally { OutlookComHost.Release(restrictedItems); }
    }

    public string CreateEvent(string subject, DateTime startDateTime, DateTime endDateTime,
        string? location, string? body, bool isMeeting, string? attendees, string? account,
        bool reminderEnabled = true, int busyStatus = OlBusy)
        => OutlookComInvoker.Run(() => CreateEventCore(subject, startDateTime, endDateTime, location, body, isMeeting, attendees, account, reminderEnabled, busyStatus));

    private string CreateEventCore(string subject, DateTime startDateTime, DateTime endDateTime,
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
        try
        {
            string stableId = (string)saved.EntryID;
            return stableId;
        }
        finally
        {
            Marshal.ReleaseComObject(saved);
            OutlookComHost.Release(ns);
        }
    }

    public bool UpdateEvent(string eventId, string? subject, DateTime? startDateTime, DateTime? endDateTime,
        string? location, string? body, string? account)
        => OutlookComInvoker.Run(() => UpdateEventCore(eventId, subject, startDateTime, endDateTime, location, body, account));

    private bool UpdateEventCore(string eventId, string? subject, DateTime? startDateTime, DateTime? endDateTime,
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
            OutlookComHost.Release(ns);
            throw new InvalidOperationException($"Event not found with ID: {eventId}");
        }

        try
        {
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
            return true;
        }
        finally
        {
            Marshal.ReleaseComObject(appointment);
            OutlookComHost.Release(ns);
        }
    }

    public bool DeleteEvent(string eventId, string? account)
        => OutlookComInvoker.Run(() => DeleteEventCore(eventId, account));

    private bool DeleteEventCore(string eventId, string? account)
    {
        var ns = GetNamespace();
        dynamic appointment;
        try
        {
            appointment = ns.GetItemFromID(eventId);
        }
        catch
        {
            OutlookComHost.Release(ns);
            throw new InvalidOperationException($"Event not found with ID: {eventId}");
        }

        if (appointment == null)
            throw new InvalidOperationException($"Event not found with ID: {eventId}");

        try
        {
            appointment.Delete();
        }
        catch (System.Runtime.InteropServices.COMException) { /* item deleted but COM reports a move error — ignore */ }
        finally
        {
            try { Marshal.ReleaseComObject(appointment); } catch { }
            OutlookComHost.Release(ns);
        }
        return true;
    }

    public List<Dictionary<string, string>> FindFreeSlots(DateTime startDate, DateTime endDate,
        int durationMinutes = 30, int workDayStart = 9, int workDayEnd = 17, string? account = null)
        => OutlookComInvoker.Run(() => FindFreeSlotsCore(startDate, endDate, durationMinutes, workDayStart, workDayEnd, account));

    private List<Dictionary<string, string>> FindFreeSlotsCore(DateTime startDate, DateTime endDate,
        int durationMinutes = 30, int workDayStart = 9, int workDayEnd = 17, string? account = null)
    {
        // Collect busy slots from all relevant calendars
        var busySlots = new List<(DateTime Start, DateTime End)>();

        void CollectBusy(dynamic folder)
        {
            var restrictedItems = GetCalendarItemsInRange(folder, startDate, endDate);
            try
            {
                dynamic? item = restrictedItems.GetFirst();
                while (item != null)
                {
                    int busyStatus = (int)item.BusyStatus;
                    if (busyStatus == OlBusy || busyStatus == OlOutOfOffice)
                        busySlots.Add(((DateTime)item.Start, (DateTime)item.End));
                    OutlookComHost.Release(item);
                    item = restrictedItems.GetNext();
                }
            }
            finally { OutlookComHost.Release(restrictedItems); }
        }

        if (string.IsNullOrEmpty(account))
        {
            var ns = GetNamespace();
            var stores = ns.Stores;
            try
            {
                for (int i = 1; i <= stores.Count; i++)
                {
                    dynamic? store = null;
                    try { store = stores.Item(i); CollectBusy(store.GetDefaultFolder(OlFolderCalendar)); }
                    catch { /* Store has no calendar folder */ }
                    finally { OutlookComHost.Release(store); }
                }
            }
            finally
            {
                OutlookComHost.Release(stores);
                OutlookComHost.Release(ns);
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
        => OutlookComInvoker.Run(() => GetAttendeeStatusCore(eventId, account));

    private Dictionary<string, object?> GetAttendeeStatusCore(string eventId, string? account)
    {
        var ns = GetNamespace();
        dynamic appointment;
        try
        {
            appointment = ns.GetItemFromID(eventId);
        }
        catch
        {
            OutlookComHost.Release(ns);
            throw new InvalidOperationException($"Event not found with ID: {eventId}");
        }

        try
        {
            if ((int)appointment.MeetingStatus != OlMeeting)
                throw new InvalidOperationException("The specified event is not a meeting.");

            var attendees = new List<Dictionary<string, string>>();
            var recipients = appointment.Recipients;
            try
            {
                for (int i = 1; i <= recipients.Count; i++)
                {
                    var recipient = recipients.Item(i);
                    try
                    {
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
                    finally { OutlookComHost.Release(recipient); }
                }
            }
            finally { OutlookComHost.Release(recipients); }

            return new Dictionary<string, object?>
            {
                ["subject"] = (string)appointment.Subject,
                ["start"] = ((DateTime)appointment.Start).ToString("yyyy-MM-dd HH:mm"),
                ["end"] = ((DateTime)appointment.End).ToString("yyyy-MM-dd HH:mm"),
                ["location"] = (string)appointment.Location,
                ["organizer"] = (string)appointment.Organizer,
                ["attendees"] = attendees
            };
        }
        finally
        {
            Marshal.ReleaseComObject(appointment);
            OutlookComHost.Release(ns);
        }
    }

    public List<Dictionary<string, object>> GetCalendars()
        => OutlookComInvoker.Run(() => GetCalendarsCore());

    private List<Dictionary<string, object>> GetCalendarsCore()
    {
        var ns = GetNamespace();
        var calendars = new List<Dictionary<string, object>>(); 

        var stores = ns.Stores;
        try
        {
            for (int i = 1; i <= stores.Count; i++)
            {
                dynamic? store = null;
                try
                {
                    store = stores.Item(i);
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
                finally { OutlookComHost.Release(store); }
            }
        }
        finally
        {
            OutlookComHost.Release(stores);
            OutlookComHost.Release(ns);
        }

        return calendars;
    }

    public List<Dictionary<string, object>> ListAccounts()
        => OutlookComInvoker.Run(() => ListAccountsCore());

    private List<Dictionary<string, object>> ListAccountsCore()
    {
        var ns = GetNamespace();
        var accounts = new List<Dictionary<string, object>>();

        var stores = ns.Stores;
        try
        {
            for (int i = 1; i <= stores.Count; i++)
            {
                dynamic? store = null;
                try
                {
                    store = stores.Item(i);
                    accounts.Add(new Dictionary<string, object>
                    {
                        ["displayName"] = (string)store.DisplayName,
                        ["storeId"] = (string)store.StoreID,
                        ["isDefault"] = i == 1
                    });
                }
                finally { OutlookComHost.Release(store); }
            }
        }
        finally
        {
            OutlookComHost.Release(stores);
            OutlookComHost.Release(ns);
        }

        return accounts;
    }

    private Dictionary<string, object?> AppointmentToDict(dynamic appointment, bool includeAttendees = false)
    {
        bool isCancelled = IsCancelledAppointment(appointment);

        var dict = new Dictionary<string, object?>
        {
            ["id"] = (string)appointment.EntryID,
            ["subject"] = (string)appointment.Subject,
            ["start"] = ((DateTime)appointment.Start).ToString("yyyy-MM-dd HH:mm"),
            ["end"] = ((DateTime)appointment.End).ToString("yyyy-MM-dd HH:mm"),
            ["location"] = (string)appointment.Location,
            ["organizer"] = (string)appointment.Organizer,
            ["isRecurring"] = (bool)appointment.IsRecurring,
            ["isMeeting"] = (int)appointment.MeetingStatus == OlMeeting,
            ["isCancelled"] = isCancelled,
            ["isAllDay"] = (bool)appointment.AllDayEvent,
            ["responseRequested"] = (bool)appointment.ResponseRequested
        };

        dict["busyStatus"] = (int)appointment.BusyStatus switch
        {
            OlBusy => "Busy",
            OlTentative => "Tentative",
            OlFree => "Free",
            OlOutOfOffice => "Out of Office",
            _ => "Unknown"
        };

        dict["responseStatus"] = (int)appointment.ResponseStatus switch
        {
            OlResponseAccepted => "Accepted",
            OlResponseDeclined => "Declined",
            OlResponseTentative => "Tentative",
            OlResponseNotResponded => "Not Responded",
            _ => "Unknown"
        };

        dict["organizerEmail"] = TryGetOrganizerEmail(appointment);
        dict["body"] = (string)appointment.Body;

        if (includeAttendees && (int)appointment.MeetingStatus is OlMeeting or OlMeetingReceived)
        {
            var attendees = new List<Dictionary<string, string>>();
            try
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
                        _ => "Not Responded"
                    };

                    attendees.Add(new Dictionary<string, string>
                    {
                        ["name"] = (string)recipient.Name,
                        ["responseStatus"] = responseStatus,
                        ["email"] = TryGetRecipientEmail(recipient)
                    });
                }
            }
            catch { /* Recipients unavailable for this item */ }
            dict["attendees"] = attendees;
        }

        // Fetch HTMLBody last — accessing it can affect COM object state
        try { dict["htmlBody"] = (string)appointment.HTMLBody; }
        catch { dict["htmlBody"] = null; }

        return dict;
    }

    private static bool IsCancelledAppointment(dynamic appointment)
    {
        try
        {
            int meetingStatus = (int)appointment.MeetingStatus;
            if (meetingStatus is OlMeetingCanceled or OlMeetingReceivedAndCanceled)
                return true;
        }
        catch { }

        try
        {
            var messageClass = ((string?)appointment.MessageClass) ?? string.Empty;
            if (messageClass.Contains("IPM.Schedule.Meeting.Canceled", StringComparison.OrdinalIgnoreCase) ||
                messageClass.Contains("IPM.Schedule.Meeting.Cancellation", StringComparison.OrdinalIgnoreCase))
                return true;
        }
        catch { }

        try
        {
            var subject = ((string?)appointment.Subject) ?? string.Empty;
            if (subject.StartsWith("Canceled:", StringComparison.OrdinalIgnoreCase) ||
                subject.StartsWith("Cancelled:", StringComparison.OrdinalIgnoreCase) ||
                subject.StartsWith("Avlyst:", StringComparison.OrdinalIgnoreCase) ||
                subject.StartsWith("Innstilt:", StringComparison.OrdinalIgnoreCase))
                return true;
        }
        catch { }

        try
        {
            var body = ((string?)appointment.Body) ?? string.Empty;
            if (body.Contains("organizer canceled this meeting", StringComparison.OrdinalIgnoreCase) ||
                body.Contains("organiser canceled this meeting", StringComparison.OrdinalIgnoreCase) ||
                body.Contains("arrangøren har avlyst dette møtet", StringComparison.OrdinalIgnoreCase) ||
                body.Contains("møtet er avlyst", StringComparison.OrdinalIgnoreCase))
                return true;
        }
        catch { }

        return false;
    }

    private static string TryGetOrganizerEmail(dynamic appointment)
    {
        try
        {
            // PR_SENT_REPRESENTING_EMAIL_ADDRESS — SMTP address of the meeting organizer
            const string PR_SENT_REPRESENTING_SMTP = "http://schemas.microsoft.com/mapi/proptag/0x0065001E";
            return (string)appointment.PropertyAccessor.GetProperty(PR_SENT_REPRESENTING_SMTP);
        }
        catch { }
        try { return (string)appointment.Organizer; }
        catch { return string.Empty; }
    }

    private static string TryGetRecipientEmail(dynamic recipient)
    {
        try
        {
            // For Exchange recipients Address is X.500; prefer SmtpAddress via PropertyAccessor
            const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
            return (string)recipient.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS);
        }
        catch { }
        try { return (string)recipient.Address; }
        catch { return string.Empty; }
    }

    /// <summary>
    /// Resolves an event ID to the appointment item to act on. Every occurrence of a recurring
    /// series shares the series' EntryID (see <c>ListEventsCore</c>/<c>AppointmentToDict</c>), so
    /// <c>GetItemFromID</c> alone always returns the recurring master — acting on it directly (e.g.
    /// responding, or editing) would apply to the whole series instead of one occurrence. When
    /// <paramref name="occurrenceStart"/> is supplied and the resolved item is in fact recurring,
    /// this looks up the specific occurrence on that date via the recurrence pattern instead.
    /// </summary>
    private static dynamic ResolveOccurrence(dynamic ns, string eventId, DateTime? occurrenceStart)
    {
        dynamic item = ns.GetItemFromID(eventId);

        if (occurrenceStart is null)
            return item;

        bool isRecurring;
        try { isRecurring = (bool)item.IsRecurring; }
        catch { isRecurring = false; }

        if (!isRecurring)
            return item;

        dynamic pattern = item.GetRecurrencePattern();
        try
        {
            // GetOccurrence matches by date only, ignoring time-of-day.
            dynamic occurrence = pattern.GetOccurrence(occurrenceStart.Value.Date);
            Marshal.ReleaseComObject(item);
            return occurrence;
        }
        catch
        {
            // No occurrence on that date (e.g. it was deleted or moved) — fall back to the
            // series master rather than failing the whole operation.
            return item;
        }
        finally { OutlookComHost.Release(pattern); }
    }

    public Dictionary<string, object?> GetEvent(string eventId, DateTime? occurrenceStart = null)
        => OutlookComInvoker.Run(() => GetEventCore(eventId, occurrenceStart));

    private Dictionary<string, object?> GetEventCore(string eventId, DateTime? occurrenceStart)
    {
        var ns = GetNamespace();
        dynamic appointment;
        try
        {
            appointment = ResolveOccurrence(ns, eventId, occurrenceStart);
        }
        catch
        {
            OutlookComHost.Release(ns);
            throw new InvalidOperationException($"Event not found with ID: {eventId}");
        }

        try
        {
            var result = AppointmentToDict(appointment, includeAttendees: true);
            return result;
        }
        finally
        {
            Marshal.ReleaseComObject(appointment);
            OutlookComHost.Release(ns);
        }
    }

    public void OpenItem(string eventId, DateTime? occurrenceStart = null)
        => OutlookComInvoker.Run(() => OpenItemCore(eventId, occurrenceStart));

    private void OpenItemCore(string eventId, DateTime? occurrenceStart)
    {
        var ns = GetNamespace();
        try
        {
            dynamic appointment = ResolveOccurrence(ns, eventId, occurrenceStart);
            try
            {
                appointment.Display(false);
                // Outlook's Inspector holds the reference — releasing here is safe
            }
            finally { Marshal.ReleaseComObject(appointment); }
        }
        catch
        {
            // Silently ignore — item may no longer exist
        }
        finally { OutlookComHost.Release(ns); }
    }

    public void RespondToMeeting(string eventId, int responseType, DateTime? occurrenceStart = null)
        => OutlookComInvoker.Run(() => RespondToMeetingCore(eventId, responseType, occurrenceStart));

    private void RespondToMeetingCore(string eventId, int responseType, DateTime? occurrenceStart)
    {
        var ns = GetNamespace();
        dynamic appointment;
        try
        {
            appointment = ResolveOccurrence(ns, eventId, occurrenceStart);
        }
        catch
        {
            OutlookComHost.Release(ns);
            throw new InvalidOperationException($"Event not found with ID: {eventId}");
        }

        try
        {
            dynamic? responseItem = appointment.Respond(responseType, true, false);
            if (responseItem != null)
            {
                try { responseItem.Send(); }
                catch { /* Response may be sent automatically with NoUI=true */ }
                finally { try { Marshal.ReleaseComObject(responseItem); } catch { } }
            }
        }
        finally
        {
            Marshal.ReleaseComObject(appointment);
            OutlookComHost.Release(ns);
        }
    }
}
