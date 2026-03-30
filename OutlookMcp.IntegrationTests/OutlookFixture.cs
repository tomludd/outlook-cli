using OutlookMcp.Services;

namespace OutlookMcp.IntegrationTests;

public class OutlookFixture : IDisposable
{
    public OutlookCalendarService CalendarService { get; }
    public OutlookMailService MailService { get; }
    public OutlookContactService ContactService { get; }

    public OutlookFixture()
    {
        CalendarService = new OutlookCalendarService();
        MailService = new OutlookMailService();
        ContactService = new OutlookContactService();
    }

    public void Dispose()
    {
        CalendarService.Dispose();
        MailService.Dispose();
        ContactService.Dispose();
        GC.SuppressFinalize(this);
    }
}
