using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;
using System.Threading.Tasks;

namespace CalendarSync
{
    public class DailySync
    {
        private readonly ITableService _tableClient;

        public DailySync(ITableService tableClient)
        {
            _tableClient = tableClient;
        }

        [FunctionName(nameof(DailySync))]
        public async Task TimerStart([TimerTrigger("0 0 3 * * *")] TimerInfo myTimer, ILogger log) // Runs every 3am UTC
        {
            var deltaLinks = _tableClient.GetDeltaLinks();
            foreach (var deltaLink in deltaLinks)
            {
                deltaLink.IsOutOfSync = true;
                await _tableClient.UpsertDeltaLink(deltaLink);
            }
        }
    }
}