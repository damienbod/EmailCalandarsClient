using Azure.Identity;
using Microsoft.Graph;
using System.Collections.Generic;
using System.Configuration;
using System.Threading.Tasks;

namespace CalendarServices.UserMailboxSettingsClient
{
    public class AadGraphApiApplicationClient
    {
        private static readonly string Tenant = ConfigurationManager.AppSettings["Tenant"];
        private static readonly string ClientId = ConfigurationManager.AppSettings["ClientId"];
        private static readonly string ClientSecret = ConfigurationManager.AppSettings["ClientSecret"];

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

        public async Task<User> GetUserMailboxSettings(string email)
        {
            var graphServiceClient = GetGraphClient();

            var userId = await GetUserIdAsync(email);

            var result = await graphServiceClient.Users[userId]
                .Request()
                .Select("MailboxSettings")
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
