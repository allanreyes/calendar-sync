using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace CalendarSync
{
    public class DetectChange
    {
        private readonly IGraphClient _graphClient;
        private readonly ITableClient _tableClient;
        private readonly IConfiguration _config;

        public DetectChange(IGraphClient graphClient, ITableClient tableClient, IConfiguration config)
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

            var calendarsToSync = new List<DeltaLink>();

            foreach (var user in users)
            {
                var hasDeltaLink = deltaLinks.Any(x => x.RowKey.Equals(user.RowKey));

                if (!hasDeltaLink)
                {
                    var deltaLink = new DeltaLink()
                    {
                        PartitionKey = "deltaLink",
                        RowKey = user.RowKey,
                        MTREmail = user.MTREmail
                    };
                    calendarsToSync.Add(deltaLink);
                }
                else
                {
                    var deltaLink = deltaLinks.Single(x => x.RowKey.Equals(user.RowKey));
                    var newDeltaLink = await _graphClient.RefreshDeltaLink(deltaLink.DeltaLinkURL);
                    var hasChanges = newDeltaLink != null;

                    if (hasChanges)
                    {
                        // update deltaLink in table storage
                        deltaLink.DeltaLinkURL = newDeltaLink;
                        await _tableClient.UpsertDeltaLink(deltaLink);
                        // Refresh calendar
                        calendarsToSync.Add(deltaLink);
                    }
                }
            }

            if (calendarsToSync.Any())
                await FindDelta(calendarsToSync, log);

            log.LogInformation($"Completed function DetectChange with ID = '{context.InvocationId}'.");
        }

        private async Task FindDelta(List<DeltaLink> calendarsToSync, ILogger log)
        {
            var tasks = new List<Task>();
            foreach (var delta in calendarsToSync)
            {
                log.LogInformation($"Adding task for {delta.MTREmail}");
                tasks.Add(SyncUserCalendarActivity(delta, log));
            }
            await Task.WhenAll(tasks);
        }

        private async Task SyncUserCalendarActivity(DeltaLink deltaLink, ILogger log)
        {
            var events = await _graphClient.GetCalendarEvents(deltaLink.RowKey);
            var mtrEvents = await _graphClient.GetCalendarEvents(deltaLink.MTREmail);

            if (events.Item2 != null)
            {
                // update refresh Token
                deltaLink.DeltaLinkURL = events.Item2;
                await _tableClient.UpsertDeltaLink(deltaLink);

                // remove events not in source calendar or was recently changed
                foreach (var @event in mtrEvents.Item1)
                    if (!events.Item1.Any(x => x.Matches(@event)))
                    {
                        log.LogInformation("Deleting Event: {0}", @event.Subject);
                        await _graphClient.DeleteCalendarEvent(deltaLink.MTREmail, @event);
                    }

                // copy events
                foreach (var @event in events.Item1)
                    if (!mtrEvents.Item1.Any(x => x.Matches(@event)))
                    {
                        log.LogInformation("Adding Event: {0}", @event.Subject);
                        await _graphClient.AddCalendarEvent(deltaLink.MTREmail, @event);
                    }
            }
        }
    }
}