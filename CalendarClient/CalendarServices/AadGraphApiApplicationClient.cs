using Azure.Identity;
using Microsoft.Graph;
using System.Collections.Generic;
using System.Configuration;
using System.Globalization;
using System.Net.Http;
using System.Threading.Tasks;

namespace CalendarServices.CalendarClient
{
    public class AadGraphApiApplicationClient
    {
        private static readonly string AadInstance = ConfigurationManager.AppSettings["AADInstance"];
        private static readonly string Tenant = ConfigurationManager.AppSettings["Tenant"];
        private static readonly string ClientId = ConfigurationManager.AppSettings["ClientId"];
        private static readonly string ClientSecret = ConfigurationManager.AppSettings["ClientSecret"];
        private static readonly string Scope = ConfigurationManager.AppSettings["Scope"];
        private static readonly string AdminUserId = ConfigurationManager.AppSettings["AdminUserId"];

        private static readonly string Authority = string.Format(CultureInfo.InvariantCulture, AadInstance, Tenant);
        private static readonly string[] Scopes = { Scope };


        private async Task<string> GetUserIdAsync(string email)
        {

            var filter = $"startswith(userPrincipalName,'{email}')";
            var graphServiceClient = GetGraphClient();

            var users = await graphServiceClient.Users
                .Request()
                .Filter(filter)
                .GetAsync();

            return users.CurrentPage[0].Id;
        }

        public async Task SendEmailAsync(Message message, string senderEmail)
        {
            var graphServiceClient = GetGraphClient();

            var saveToSentItems = true;

            var userId = await GetUserIdAsync(senderEmail);

            await graphServiceClient.Users[userId]
                .SendMail(message, saveToSentItems)
                .Request()
                .PostAsync();
        }

        public async Task<IUserCalendarViewCollectionPage> GetCalanderForUser(string email, string from, string to)
        {
            var graphServiceClient = GetGraphClient();

            var userId = await GetUserIdAsync(email);
            // var filter = "startsWith(subject,'All')";
            var queryOptions = new List<QueryOption>()
            {
                new QueryOption("startDateTime", from),
                new QueryOption("endDateTime", to)
            };

            //var result = await graphServiceClient.Users[userId].Calendar.Events
            var result = await graphServiceClient.Users[userId].CalendarView
                .Request(queryOptions)
                .Select("start,end,subject,location,sensitivity, showAs, isAllDay")
                // .Filter(filter)
                .GetAsync();

            return result;
        }

        private GraphServiceClient GetGraphClient()
        {
            string[] scopes = new[] { "https://graph.microsoft.com/.default" };
            var tenantId = Tenant;
            var clientId = ClientId;
            var clientSecret = ClientSecret;

            var options = new TokenCredentialOptions
            {
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
            };

            // https://docs.microsoft.com/dotnet/api/azure.identity.clientsecretcredential
            var clientSecretCredential = new ClientSecretCredential(
                tenantId, clientId, clientSecret, options);

            return new GraphServiceClient(clientSecretCredential, scopes);
        }
    }
}
