using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;
using System.Threading.Tasks;

namespace CalendarSync
{
    public class FindUsersWithMTR
    {
        private readonly IGraphClient _graphClient;
        private readonly ITableService _tableService;

        private readonly string _prefix;

        public FindUsersWithMTR(IGraphClient graphClient, ITableService tableService)
        {
            _graphClient = graphClient;
            _tableService = tableService;
        }

        [FunctionName(nameof(FindUsersWithMTR))]
        public async Task Run([TimerTrigger("0 0 3 * * *", RunOnStartup = true)] TimerInfo myTimer, ILogger log)
        {
            await _tableService.TruncateUsersTable();
            var users = await _graphClient.GetMTRUsers();
            await _tableService.AddUsers(users);
        }
    }
}