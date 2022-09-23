using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CalendarSync
{
    public class PerformChange
    {
        private readonly IGraphClient _graphClient;
        private readonly ITableService _tableService;
        private readonly IConfiguration _config;

        public PerformChange(IGraphClient graphClient, ITableService tableService, IConfiguration config)
        {
            _graphClient = graphClient;
            _tableService = tableService;
            _config = config;
        }

        [FunctionName(nameof(PerformChange))]
        public async Task TimerStart(
            [TimerTrigger("*/5 * * * * *")] TimerInfo myTimer,
            ExecutionContext context,
            ILogger log)
        {
            var deltaLinks = _tableService.GetDeltaLinks();
            foreach (var deltaLink in deltaLinks.Where(d => d.IsOutOfSync))
            {
                var newDeltaLink = await SyncUserCalendarActivity(deltaLink);

                // Revert back to in sync status
                deltaLink.DeltaLinkURL = newDeltaLink;
                deltaLink.IsOutOfSync = false;
                await _tableService.UpsertDeltaLink(deltaLink);
            }
        }     

        private async Task<string> SyncUserCalendarActivity(DeltaLink deltaLink)
        {
            var user = _tableService.GetUsers().SingleOrDefault(p=> p.RowKey.Equals(deltaLink.RowKey, StringComparison.InvariantCultureIgnoreCase));
            if (user == null) return deltaLink.DeltaLinkURL; // User does not have MTR account. Might have in the past but now deleted.

            var events = await _graphClient.GetCalendarEvents(deltaLink.RowKey);
            var mtrEvents = await _graphClient.GetCalendarEvents(deltaLink.MTREmail);
            
            var syncActions = new StringBuilder();

            if (events.Item2 != null)
            {
                // update refresh Token
                deltaLink.DeltaLinkURL = events.Item2;
                await _tableService.UpsertDeltaLink(deltaLink);

                // remove events not in source calendar or was recently changed
                foreach (var @event in mtrEvents.Item1)
                    if (!events.Item1.Any(x => x.Matches(@event)))
                    {
                        await _graphClient.DeleteCalendarEvent(deltaLink.MTREmail, @event);
                        syncActions.AppendLine($"Deleted: {@event.Subject}");
                    }

                // copy events
                foreach (var @event in events.Item1)
                    if (!mtrEvents.Item1.Any(x => x.Matches(@event)))
                    {
                        var result = await _graphClient.AddCalendarEvent(deltaLink.MTREmail, @event);
                        if (result == null)
                        {
                            syncActions.AppendLine($"Added: {@event.Subject}");
                        }
                        else {
                            syncActions.AppendLine($"Not Added: {result}");
                        }
                    }

                await _tableService.LogSyncActions(deltaLink.RowKey, syncActions.ToString());
            }

            return events.Item2; // new deltaLink url
        }
    }
}