Imports Microsoft.Graph
Imports Microsoft.Identity.Client
Imports System.Configuration
Imports System.Globalization
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports VbEmailClient.GraphEmailClient

Namespace EmailCalendarsClient.MailSender
    Public Class AadGraphApiDelegatedClient
        Private ReadOnly _httpClient As HttpClient = New HttpClient()
        Private _app As IPublicClientApplication
        Private Shared ReadOnly AadInstance As String = ConfigurationManager.AppSettings("AADInstance")
        Private Shared ReadOnly Tenant As String = ConfigurationManager.AppSettings("Tenant")
        Private Shared ReadOnly ClientId As String = ConfigurationManager.AppSettings("ClientId")
        Private Shared ReadOnly Scope As String = ConfigurationManager.AppSettings("Scope")
        Private Shared ReadOnly Authority As String = String.Format(CultureInfo.InvariantCulture, AadInstance, Tenant)
        Private Shared ReadOnly Scopes As String() = {Scope}

        Public Sub InitClient()
            _app = PublicClientApplicationBuilder.Create(ClientId).WithAuthority(Authority).WithRedirectUri("http://localhost:65419").Build()
            TokenCacheHelper.EnableSerialization(_app.UserTokenCache)
        End Sub

        Public Async Function SignIn() As Task(Of IAccount)
            Dim accounts = (Await _app.GetAccountsAsync()).ToList()

            Try
                Dim result = Await _app.AcquireTokenSilent(Scopes, accounts.FirstOrDefault()).ExecuteAsync().ConfigureAwait(False)
                Return result.Account
            Catch __unusedMsalUiRequiredException1__ As MsalUiRequiredException
                Dim builder = _app.AcquireTokenInteractive(Scopes).WithAccount(accounts.FirstOrDefault()).WithUseEmbeddedWebView(False).WithPrompt(Microsoft.Identity.Client.Prompt.SelectAccount)
                Dim result = builder.ExecuteAsync().GetAwaiter().GetResult()
                Return result.Account
            End Try
        End Function

        Public Async Function AcquireTokenSilent() As Task(Of AuthenticationResult)
            Dim accounts = Await GetAccountsAsync()
            Dim result = Await _app.AcquireTokenSilent(Scopes, accounts.FirstOrDefault()).ExecuteAsync().ConfigureAwait(False)
            Return result
        End Function

        Public Async Function GetAccountsAsync() As Task(Of IList(Of IAccount))
            Dim accounts = Await _app.GetAccountsAsync()
            Return accounts.ToList()
        End Function

        Public Async Function RemoveAccountsAsync() As Task
            Dim accounts As IList(Of IAccount) = Await GetAccountsAsync()

            While accounts.Any()
                Await _app.RemoveAsync(accounts.First())
                accounts = Await GetAccountsAsync()
            End While
        End Function

        Public Async Function SendEmailAsync(ByVal message As Message) As Task
            Dim result = Await AcquireTokenSilent()
            _httpClient.DefaultRequestHeaders.Authorization = New AuthenticationHeaderValue("Bearer", result.AccessToken)
            _httpClient.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/json"))
            Dim graphClient As GraphServiceClient = New GraphServiceClient(_httpClient) With {
                .AuthenticationProvider = New DelegateAuthenticationProvider(Async Function(requestMessage)
                                                                                 requestMessage.Headers.Authorization = New AuthenticationHeaderValue("Bearer", result.AccessToken)
                                                                             End Function)
            }
            Dim saveToSentItems = True
            Await graphClient.[Me].SendMail(message, saveToSentItems).Request().PostAsync()
        End Function
    End Class
End Namespace
