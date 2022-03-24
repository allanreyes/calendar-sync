using Microsoft.Azure.Cosmos.Table;

namespace CalendarSync
{
    public class UsersWithMTR : TableEntity
    {
        public string MTREmail { get; set; }
        public string TimeZone { get; set; }
    }
}
