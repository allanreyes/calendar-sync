using Microsoft.Azure.Cosmos.Table;
using Microsoft.Extensions.Configuration;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace CalendarSync
{
    public class TableClient : ITableClient
    {
        private readonly CloudTable _usersTable;
        private readonly CloudTable _deltaLinksTable;
        public TableClient(IConfiguration config)
        {
            var account = CloudStorageAccount.Parse(config["AzureWebJobsStorage"]);
            var client = account.CreateCloudTableClient();
            _usersTable = client.GetTableReference("UsersWithMTR");
            _usersTable.CreateIfNotExists();
            _deltaLinksTable = client.GetTableReference("DeltaLinks");
            _deltaLinksTable.CreateIfNotExists();
        }

        #region UsersWithMTR

        public IEnumerable<UsersWithMTR> GetUsers()
        {
            return _usersTable.ExecuteQuery(new TableQuery<UsersWithMTR>()).ToList();
        }

        public async Task TruncateUsersTable()
        {
            var users = GetUsers();
            var allTasks = users.Select(async r =>
            {
                await _usersTable.ExecuteAsync(TableOperation.Delete(r));
            });
            await Task.WhenAll(allTasks);
        }

        public async Task AddUsers(IEnumerable<UsersWithMTR> users)
        {
            var batch = new TableBatchOperation();
            foreach (var user in users)
            {
                batch.Insert(user);
            }
            await _usersTable.ExecuteBatchAsync(batch);
        }

        #endregion

        #region DeltaLinks

        public IEnumerable<DeltaLink> GetDeltaLinks()
        {
            return _deltaLinksTable.ExecuteQuery(new TableQuery<DeltaLink>()).ToList();
        }

        public async Task UpsertDeltaLink(DeltaLink deltaLink)
        {
            var operation = TableOperation.InsertOrMerge(deltaLink);
            await _deltaLinksTable.ExecuteAsync(operation);
        }

        #endregion
    }

    public interface ITableClient
    {
        Task UpsertDeltaLink(DeltaLink deltaLink);
        Task AddUsers(IEnumerable<UsersWithMTR> users);
        IEnumerable<DeltaLink> GetDeltaLinks();
        IEnumerable<UsersWithMTR> GetUsers();
        Task TruncateUsersTable();

    }
}
