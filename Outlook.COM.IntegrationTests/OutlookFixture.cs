using Outlook.COM;

namespace Outlook.COM.IntegrationTests;

public class OutlookFixture
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
}
