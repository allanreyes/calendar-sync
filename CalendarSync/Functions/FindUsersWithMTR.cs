using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace CalendarSync
{
    public class FindUsersWithMTR
    {
        private readonly IGraphClient _graphClient;
        private readonly ITableClient _tableClient;

        private readonly string _prefix;

        public FindUsersWithMTR(IGraphClient graphClient, ITableClient tableClient, IConfiguration config)
        {
            _graphClient = graphClient;
            _tableClient = tableClient;
            _prefix = config["Prefix"];
        }

        [FunctionName(nameof(FindUsersWithMTR))]
        public async Task Run([TimerTrigger("0 */10 * * * *", RunOnStartup = true)] TimerInfo myTimer, ILogger log)
        {
            await _tableClient.TruncateUsersTable();

            var mtrAccounts = await _graphClient.GetMTRAccounts();
            var users = new List<UsersWithMTR>();
            var partitionKey = Guid.NewGuid().ToString();

            foreach (var mtrAccount in mtrAccounts)
            {
                var userEmail = mtrAccount.Mail.Substring(_prefix.Length);
                var userAccount = await _graphClient.GetUser(userEmail);

                if (userAccount != null)
                {
                    users.Add(new UsersWithMTR()
                    {
                        PartitionKey = partitionKey,
                        RowKey = userAccount.Mail,
                        MTREmail = mtrAccount.Mail,
                        TimeZone = userAccount.MailboxSettings.TimeZone
                    });
                }
            }

            await _tableClient.AddUsers(users);
        }
    }
}
