using Microsoft.Graph;
using Microsoft.Identity.Client;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Configuration;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace CalendarServices.UserMailboxSettingsClient
{
    /// <summary>
    /// Second way of create a GraphServiceClient
    /// </summary>
    public class ApiTokenInMemoryClient
    {
        private readonly HttpClient _httpClient = new HttpClient();
        private static readonly string Tenant = ConfigurationManager.AppSettings["Tenant"];
        private static readonly string ClientId = ConfigurationManager.AppSettings["ClientId"];
        private static readonly string ClientSecret = ConfigurationManager.AppSettings["ClientSecret"];

        private readonly IConfidentialClientApplication _app;
        private readonly ConcurrentDictionary<string, AccessTokenItem> _accessTokens = new();

        private class AccessTokenItem
        {
            public string AccessToken { get; set; } = string.Empty;
            public DateTime ExpiresIn { get; set; }
        }

        public ApiTokenInMemoryClient()
        {
            _app = InitConfidentialClientApplication();
        }

        public async Task<GraphServiceClient> GetGraphClient()
        {
            var result = await GetApiToken("default");

            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", result);
            _httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            var graphClient = new GraphServiceClient(_httpClient)
            {
                AuthenticationProvider = new DelegateAuthenticationProvider(async (requestMessage) =>
                {
                    requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", result);
                    await Task.FromResult<object>(null);
                })
            };

            return graphClient;
        }

        private async Task<string> GetApiToken(string api_name)
        {
            if (_accessTokens.ContainsKey(api_name))
            {
                var accessToken = _accessTokens.GetValueOrDefault(api_name);
                if (accessToken.ExpiresIn > DateTime.UtcNow)
                {
                    return accessToken.AccessToken;
                }
                else
                {
                    // remove
                    _accessTokens.TryRemove(api_name, out _);
                }
            }

            // add
            var newAccessToken = await AcquireTokenSilent();
            _accessTokens.TryAdd(api_name, newAccessToken);

            return newAccessToken.AccessToken;
        }

        private async Task<AccessTokenItem> AcquireTokenSilent()
        {
            //var scopes = "User.read Mail.Send Mail.ReadWrite OnlineMeetings.ReadWrite.All";
            var authResult = await _app
                .AcquireTokenForClient(scopes: new[] { "https://graph.microsoft.com/.default" })
                .WithAuthority(AzureCloudInstance.AzurePublic, Tenant)
                .ExecuteAsync();

            return new AccessTokenItem
            {
                ExpiresIn = authResult.ExpiresOn.UtcDateTime,
                AccessToken = authResult.AccessToken
            };
        }

        private IConfidentialClientApplication InitConfidentialClientApplication()
        {
            return ConfidentialClientApplicationBuilder
                .Create(ClientId)
                .WithClientSecret(ClientSecret)
                .Build();
        }
    }
}
