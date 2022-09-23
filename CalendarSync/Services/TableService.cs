using Azure.Data.Tables;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace CalendarSync
{
    public class TableService : ITableService
    {
        private readonly TableClient _usersTable;
        private readonly TableClient _deltaLinksTable;
        private readonly TableClient _deltaLinksLogTable;
        private readonly TableClient _syncActionsTable;

        public TableService(IConfiguration config)
        {
            _usersTable = new TableClient(config["AzureWebJobsStorage"], "UsersWithMTR");
            _usersTable.CreateIfNotExists();
            _deltaLinksTable = new TableClient(config["AzureWebJobsStorage"], "DeltaLinks");
            _deltaLinksTable.CreateIfNotExists();
            _deltaLinksLogTable = new TableClient(config["AzureWebJobsStorage"], "DeltaLinksLog");
            _deltaLinksLogTable.CreateIfNotExists();
            _syncActionsTable = new TableClient(config["AzureWebJobsStorage"], "SyncActionsLog");
            _syncActionsTable.CreateIfNotExists();
        }

        #region UsersWithMTR

        public IEnumerable<UsersWithMTR> GetUsers()
        {
            return _usersTable.Query<UsersWithMTR>().ToList();
        }

        public async Task TruncateUsersTable()
        {
            var users = GetUsers();
            var allTasks = users.Select(async r =>
            {
                await _usersTable.DeleteEntityAsync(r.PartitionKey, r.RowKey);
            });
            await Task.WhenAll(allTasks);
        }

        public async Task AddUsers(IEnumerable<UsersWithMTR> users)
        {
            var batch = new List<TableTransactionAction>();
            batch.AddRange(users.Select(e => new TableTransactionAction(TableTransactionActionType.Add, e)));
            await _usersTable.SubmitTransactionAsync(batch).ConfigureAwait(false);
        }

        #endregion

        #region DeltaLinks

        public IEnumerable<DeltaLink> GetDeltaLinks()
        {
            return _deltaLinksTable.Query<DeltaLink>().ToList();
        }

        public async Task UpsertDeltaLink(DeltaLink deltaLink)
        {
            await _deltaLinksTable.UpsertEntityAsync(deltaLink);
        }

        public async Task LogDeltaLinkChange(string emailAddress, string @events)
        {
            var entity = new TableEntity(emailAddress, Guid.NewGuid().ToString()){
                { "Events", @events}
            };
            await _deltaLinksLogTable.AddEntityAsync(entity);
        }

        public async Task LogSyncActions(string emailAddress, string actions)
        {
            var entity = new TableEntity(emailAddress, Guid.NewGuid().ToString()){
                { "Actions", actions}
            };
            await _syncActionsTable.AddEntityAsync(entity);
        }
        #endregion
    }

    public interface ITableService
    {
        Task UpsertDeltaLink(DeltaLink deltaLink);
        Task LogDeltaLinkChange(string email, string emailSubjects);
        Task LogSyncActions(string email, string emailSubjects);

        Task AddUsers(IEnumerable<UsersWithMTR> users);
        IEnumerable<DeltaLink> GetDeltaLinks();
        IEnumerable<UsersWithMTR> GetUsers();
        Task TruncateUsersTable();

    }
}
