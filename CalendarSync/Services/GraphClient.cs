using Azure.Core;
using Azure.Identity;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Graph.TermStore;
using Newtonsoft.Json;
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
        private readonly ITableService _tableService;

        public GraphClient(IConfiguration config, ILogger<GraphClient> logger, ITableService tableService)
        {
            _config = config;
            _prefix = config["Prefix"];
            _logger = logger;
            _tableService = tableService;

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
                return await _graphClient.Users[email].Request()?.GetAsync();
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
            var queryStartDate = DateTime.Now.AddDays(Convert.ToInt32(_config["DaysAgo"]));
            var queryEndDate = DateTime.Now.AddDays(Convert.ToInt32(_config["DaysFromToday"]));

            var options = new List<QueryOption>()
                {
                    new QueryOption("startDateTime", queryStartDate.ToString("o")),
                    new QueryOption("endDateTime", queryEndDate.ToString("o"))
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

        public async Task<string> AddCalendarEvent(string email, Event calendarEvent)
        {
            if (!email.StartsWith(_prefix, StringComparison.InvariantCultureIgnoreCase))
                return $"Account {email} is not an MTR account";

            if (calendarEvent?.Subject == null)
                return $"(Empty Subject)";

            if (calendarEvent.Attendees.Count() == 0 && calendarEvent?.IsOrganizer == true)
                return $"{calendarEvent.Subject} (Appointment)";

            if (calendarEvent.Sensitivity == Sensitivity.Private)
                return $"{calendarEvent.Subject} (Private)";

            if (calendarEvent?.IsOnlineMeeting == false)
                return $"{calendarEvent?.Subject} (Not an Online meeting)";

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
                        DateTime = TimeZoneInfo.ConvertTimeFromUtc(
                            DateTime.Parse(calendarEvent.Start.DateTime),
                            TimeZoneInfo.FindSystemTimeZoneById(calendarEvent.OriginalStartTimeZone)).ToString(),
                        TimeZone = calendarEvent.OriginalStartTimeZone
                    },
                    End = new DateTimeTimeZone
                    {
                        DateTime = TimeZoneInfo.ConvertTimeFromUtc(
                            DateTime.Parse(calendarEvent.End.DateTime),
                            TimeZoneInfo.FindSystemTimeZoneById(calendarEvent.OriginalEndTimeZone)).ToString(),
                        TimeZone = calendarEvent.OriginalEndTimeZone
                    },
                    Location = new Location(),
                    Attendees = new List<Attendee>(),
                    Recurrence = calendarEvent.Recurrence
                };

                await _graphClient.Users[email].Calendar.Events.Request().AddAsync(@event);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, ex.Message);
            }
            return null;
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

        public async Task<DeltaLink> RefreshDeltaLink(DeltaLink deltaLink)
        {
            try
            {
                HttpRequestMessage hrm = new HttpRequestMessage(HttpMethod.Get, deltaLink.DeltaLinkURL);
                await _graphClient.AuthenticationProvider.AuthenticateRequestAsync(hrm);
                HttpResponseMessage response = await _graphClient.HttpProvider.SendAsync(hrm);
                var responseString = await response.Content.ReadAsStringAsync();
                var data = JObject.Parse(responseString);
                deltaLink.DeltaLinkURL = data["@odata.deltaLink"].ToString();

                var deltaEvents = (JArray)data["value"];
                if (deltaEvents.Any()) // Has changes
                {
                    deltaLink.IsOutOfSync = true;
                    var eventDetails = deltaEvents.Select(e => new { Subject = e["subject"], Organizer = e["organizer"] });
                    await _tableService.LogDeltaLinkChange(deltaLink.RowKey, JsonConvert.SerializeObject(eventDetails));
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, ex.Message);
            }
            return deltaLink;
        }
    }

    public interface IGraphClient
    {
        Task<string> AddCalendarEvent(string email, Event calendarEvent);
        Task DeleteCalendarEvent(string email, Event eventToDelete);
        Task<(IEnumerable<Event>, string)> GetCalendarEvents(string userId);
        Task<IEnumerable<User>> GetMTRAccounts();
        Task<User> GetUser(string email);
        Task<DeltaLink> RefreshDeltaLink(DeltaLink deltaLink);
    }
}
