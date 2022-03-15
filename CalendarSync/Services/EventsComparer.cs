using Microsoft.Graph;

namespace CalendarSync
{
    public static class EventComparer
    {
        public static bool Matches(this Event x, Event y)
        {
            return x.Subject == y.Subject &&
                   x.Body.ContentType == y.Body.ContentType &&
                   x.Start.DateTime == y.Start.DateTime &&
                   x.Start.TimeZone == y.Start.TimeZone &&
                   x.End.DateTime == y.End.DateTime &&
                   x.End.TimeZone == y.End.TimeZone;
        }
    }
}
