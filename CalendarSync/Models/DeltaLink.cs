using Azure;
using Azure.Data.Tables;
using System;

namespace CalendarSync
{
    public class DeltaLink : ITableEntity
    {
        public string MTREmail { get; set; }
        
        public bool IsOutOfSync { get; set; }

        public string DeltaLinkURL { get; set; }
        public string PartitionKey { get; set; }
        public string RowKey { get; set; }
        public DateTimeOffset? Timestamp { get; set; }
        public ETag ETag { get; set; }
    }
}
