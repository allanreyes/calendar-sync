using Microsoft.Azure.Cosmos.Table;

namespace CalendarSync
{
    public class DeltaLink : TableEntity
    {
        public string MTREmail { get; set; }

        public string DeltaLinkURL { get; set; }
    }
}
