

using Microsoft.Graph;
using Microsoft.Identity.Client;
using System.Collections.Generic;
using System.Configuration;
using System.Globalization;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using TodoListClient;

namespace EmailCalendarsClient.MailSender
{
    public class AadNativeClient
    {
        private readonly HttpClient _httpClient = new HttpClient();
        private IPublicClientApplication _app;

        private static readonly string AadInstance = ConfigurationManager.AppSettings["ida:AADInstance"];
        private static readonly string Tenant = ConfigurationManager.AppSettings["ida:Tenant"];
        private static readonly string ClientId = ConfigurationManager.AppSettings["ida:ClientId"];

        private static readonly string Authority = string.Format(CultureInfo.InvariantCulture, AadInstance, Tenant);

        private static readonly string Scope = ConfigurationManager.AppSettings["todo:Scope"];
        private static readonly string[] Scopes = { Scope };

        public void InitClient()
        {
            _app = PublicClientApplicationBuilder.Create(ClientId)
                .WithAuthority(Authority)
                .WithRedirectUri("http://localhost:65419") // needed only for the system browser
                .Build();

            TokenCacheHelper.EnableSerialization(_app.UserTokenCache);
        }

        public async Task<IAccount> SignIn()
        {
            var accounts = (await _app.GetAccountsAsync()).ToList();
            try
            {
                var result = await _app.AcquireTokenSilent(Scopes, accounts.FirstOrDefault())
                    .ExecuteAsync()
                    .ConfigureAwait(false);

                return result.Account;
            }
            catch (MsalUiRequiredException)
            {
                var builder = _app.AcquireTokenInteractive(Scopes)
                    .WithAccount(accounts.FirstOrDefault())
                    .WithUseEmbeddedWebView(false)
                    .WithPrompt(Microsoft.Identity.Client.Prompt.SelectAccount);

                var result = await builder.ExecuteAsync().ConfigureAwait(false);

                return result.Account;
            }
        }

        public async Task<AuthenticationResult> AcquireTokenSilent()
        {
            var accounts = await GetAccountsAsync();
            var result = await _app.AcquireTokenSilent(Scopes, accounts.FirstOrDefault())
                    .ExecuteAsync()
                    .ConfigureAwait(false);

            return result;
        }          

        public async Task<IList<IAccount>> GetAccountsAsync()
        {
            var accounts = await _app.GetAccountsAsync();
            return accounts.ToList();
        }

        public async Task RemoveAccountsAsync()
        {
            IList<IAccount> accounts = await GetAccountsAsync();

            // Clears the library cache. Does not affect the browser cookies.
            while (accounts.Any())
            {
                await _app.RemoveAsync(accounts.First());
                accounts = await GetAccountsAsync();
            }
        }

        public async Task SendMailAsync()
        {
            var result = await AcquireTokenSilent();

            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", result.AccessToken);
            _httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            GraphServiceClient graphClient = new GraphServiceClient(_httpClient)
            {
                AuthenticationProvider = new DelegateAuthenticationProvider(async (requestMessage) =>
                {
                    requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", result.AccessToken);
                })
            };


            var message = new Message
            {
                Subject = "Meet for lunch?",
                Body = new ItemBody
                {
                    ContentType = BodyType.Text,
                    Content = "The new cafeteria is open."
                },
                ToRecipients = new List<Recipient>()
                {
                    new Recipient
                    {
                        EmailAddress = new EmailAddress
                        {
                            Address = "damien_bod@hotmail.com"
                        }
                    }
                },
                CcRecipients = new List<Recipient>()
                {
                    new Recipient
                    {
                        EmailAddress = new EmailAddress
                        {
                            Address = "danas@contoso.onmicrosoft.com"
                        }
                    }
                }
            };

            var saveToSentItems = true;

            await graphClient.Me
                .SendMail(message, saveToSentItems)
                .Request()
                .PostAsync();
        }

    }
}
