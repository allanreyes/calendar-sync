using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System.Linq;
using System.Threading.Tasks;

namespace CalendarSync
{
    public class DetectChange
    {
        private readonly IGraphClient _graphClient;
        private readonly ITableService _tableClient;
        private readonly IConfiguration _config;

        public DetectChange(IGraphClient graphClient, ITableService tableClient, IConfiguration config)
        {
            _graphClient = graphClient;
            _tableClient = tableClient;
            _config = config;
        }

        [FunctionName(nameof(DetectChange))]
        public async Task TimerStart(
            [TimerTrigger("*/30 * * * * *")] TimerInfo myTimer,
            ExecutionContext context,
            ILogger log)
        {
            log.LogInformation($"Started function DetectChange with ID = '{context.InvocationId}'.");

            var users = _tableClient.GetUsers();
            var deltaLinks = _tableClient.GetDeltaLinks();

            foreach (var user in users)
            {
                var deltaLink = deltaLinks.SingleOrDefault(x => x.RowKey.Equals(user.RowKey));

                if (deltaLink == null) // This is the first time this email is being onboarded so no Delta Link yet
                {
                    var calendarEvents = await _graphClient.GetCalendarEvents(user.RowKey);
                    deltaLink = new DeltaLink()
                    {
                        PartitionKey = "deltaLink",
                        RowKey = user.RowKey,
                        MTREmail = user.MTREmail,
                        IsOutOfSync = true,
                        DeltaLinkURL = calendarEvents.Item2
                    };
                }
                else
                {
                    deltaLink = await _graphClient.RefreshDeltaLink(deltaLink);
                }
                await _tableClient.UpsertDeltaLink(deltaLink);
            }

            log.LogInformation($"Completed function DetectChange with ID = '{context.InvocationId}'.");
        }
    }
}