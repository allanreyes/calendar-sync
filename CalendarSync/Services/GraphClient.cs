using Azure.Core;
using Azure.Identity;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace CalendarSync
{
    public class GraphClient : IGraphClient
    {
        private readonly IConfiguration _config;
        private readonly ILogger _logger;
        private readonly string _prefix;
        private readonly GraphServiceClient _graphClient;
        private readonly HttpClient _httpClient;

        public GraphClient(IConfiguration config, ILogger<GraphClient> logger)
        {
            _config = config;
            _prefix = config["Prefix"];
            _logger = logger;

            var scopes = new[] { "https://graph.microsoft.com/.default" };

            var options = new TokenCredentialOptions
            {
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
            };
            var clientSecretCredential = new ClientSecretCredential(
                config["TenantId"], config["ClientId"], config["ClientSecret"], options);
            _graphClient = new GraphServiceClient(clientSecretCredential, scopes);

            var ctx = new TokenRequestContext(scopes: new[] { "https://graph.microsoft.com/.default" });
            var token = clientSecretCredential.GetToken(ctx).Token;
            _httpClient = new HttpClient();
            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
        }

        public async Task<IEnumerable<User>> GetMTRAccounts()
        {
            _logger.LogInformation($"Finding email address that start with '{_prefix}'");
            try
            {
                return await _graphClient.Users.Request()
                                            .Filter($"startswith(mail, '{_prefix}')")
                                            .Select(x => new { x.DisplayName, x.Id, x.Mail })
                                            .GetAsync();
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, ex.Message);
                return Enumerable.Empty<User>();
            }
        }

        public async Task<User> GetUser(string email)
        {
            _logger.LogInformation($"Getting user '{email}'");
            try
            {
                var user = await _graphClient.Users[email].Request()?.GetAsync();
                var response = await _httpClient.GetAsync($"https://graph.microsoft.com/v1.0/users/{user.Id}/mailboxSettings");
                var mailboxSettingsJson = await response.Content.ReadAsStringAsync();
                var mailboxSettings = JObject.Parse(mailboxSettingsJson);
                user.MailboxSettings = new MailboxSettings() { TimeZone = mailboxSettings.GetValue("timeZone").ToString() };
                return user;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, ex.Message);
                return null;
            }
        }

        public async Task<(IEnumerable<Event>, string)> GetCalendarEvents(string userId)
        {
            _logger.LogInformation($"Getting calendar events of '{userId}'");

            var options = new List<QueryOption>()
                {
                    new QueryOption("startDateTime", DateTime.Now.AddDays(Convert.ToInt32(_config["DaysAgo"])).ToString("o")),
                    new QueryOption("endDateTime", DateTime.Now.AddDays(Convert.ToInt32(_config["DaysFromToday"])).ToString("o"))
                };

            var page = await _graphClient.Users[userId].CalendarView.Delta().Request(options).GetAsync();
            var result = new List<Event>();
            string deltaLink = null;

            if (page != null)
            {
                do
                {
                    foreach (var calendarEvent in page.CurrentPage)
                        if (calendarEvent is Event)
                            result.Add(calendarEvent);

                    if (page.NextPageRequest != null)
                    {
                        page = await page.NextPageRequest.GetAsync();
                    }
                    else
                    {
                        deltaLink = page.AdditionalData["@odata.deltaLink"].ToString();
                        break;
                    }
                } while (true);
            }

            return (result, deltaLink);
        }

        public async Task AddCalendarEvent(string email, string timeZone, Event calendarEvent)
        {
            var isMTRAccount = email.StartsWith(_prefix, StringComparison.InvariantCultureIgnoreCase);
            var isAppointnment = !calendarEvent.Attendees.Any();
            var isPrivate = calendarEvent.Sensitivity == Sensitivity.Private;
            var isEmptySubject = calendarEvent?.Subject == null;

            if (isMTRAccount && !isAppointnment && !isPrivate && !isEmptySubject)
            {
                try
                {
                    var @event = new Event
                    {
                        Subject = calendarEvent.Subject,
                        Body = new ItemBody
                        {
                            ContentType = calendarEvent.Body.ContentType,
                            Content = calendarEvent.Body.Content
                        },
                        Start = new DateTimeTimeZone
                        {
                            DateTime = calendarEvent.Start.DateTime + "Z",
                            TimeZone = timeZone

                        },
                        End = new DateTimeTimeZone
                        {
                            DateTime = calendarEvent.End.DateTime + "Z",
                            TimeZone = timeZone
                        },
                        Location = new Location(),
                        Attendees = new List<Attendee>()
                    };

                    await _graphClient.Users[email].Calendar.Events.Request().AddAsync(@event);
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, ex.Message);
                }
            }
        }

        public async Task DeleteCalendarEvent(string email, Event eventToDelete)
        {
            // Make sure calendar manipulation is done only against the MTR account, and not user accounts
            if (!email.StartsWith(_prefix, StringComparison.InvariantCultureIgnoreCase))
                return;

            try
            {
                await _graphClient.Users[email].Events[eventToDelete.Id].Request().DeleteAsync();
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, ex.Message);
            }
        }

        public async Task<string> RefreshDeltaLink(string deltaLink)
        {
            try
            {
                HttpRequestMessage hrm = new HttpRequestMessage(HttpMethod.Get, deltaLink);
                await _graphClient.AuthenticationProvider.AuthenticateRequestAsync(hrm);
                HttpResponseMessage response = await _graphClient.HttpProvider.SendAsync(hrm);
                var responseString = await response.Content.ReadAsStringAsync();
                var data = JObject.Parse(responseString);
                var hasChanges = ((JArray)data["value"]).Count > 0;
                if (hasChanges)
                {
                    return data["@odata.deltaLink"].ToString();
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, ex.Message);
            }
            return null;
        }
    }

    public interface IGraphClient
    {
        Task AddCalendarEvent(string email, string timeZone, Event calendarEvent);
        Task DeleteCalendarEvent(string email, Event eventToDelete);
        Task<(IEnumerable<Event>, string)> GetCalendarEvents(string userId);
        Task<IEnumerable<User>> GetMTRAccounts();
        Task<User> GetUser(string email);
        Task<string> RefreshDeltaLink(string deltaLink);
    }
}
