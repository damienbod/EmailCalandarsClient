Imports Microsoft.Graph
Imports Microsoft.Identity.Client
Imports System.Configuration
Imports System.Globalization
Imports System.Net.Http
Imports System.Net.Http.Headers

Namespace MailboxSettings
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
            Try
                Dim result = Await AcquireTokenSilent()
                Return result.Account
            Catch __unusedMsalUiRequiredException1__ As MsalUiRequiredException
                Return AcquireTokenInteractive().GetAwaiter().GetResult()
            End Try
        End Function

        Private Async Function AcquireTokenInteractive() As Task(Of IAccount)
            Dim accounts = (Await _app.GetAccountsAsync()).ToList()
            Dim builder = _app.AcquireTokenInteractive(Scopes).WithAccount(accounts.FirstOrDefault()).WithUseEmbeddedWebView(False).WithPrompt(Microsoft.Identity.Client.Prompt.SelectAccount)
            Dim result = Await builder.ExecuteAsync().ConfigureAwait(False)
            Return result.Account
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

        Public Async Function GetUserMailboxSettings(ByVal email As String) As Task(Of User)
            Dim result = Await AcquireTokenSilent()
            _httpClient.DefaultRequestHeaders.Authorization = New AuthenticationHeaderValue("Bearer", result.AccessToken)
            _httpClient.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/json"))
            Dim graphClient = New GraphServiceClient(_httpClient) With {
                .AuthenticationProvider = New DelegateAuthenticationProvider(Async Function(requestMessage)
                                                                                 requestMessage.Headers.Authorization = New AuthenticationHeaderValue("Bearer", result.AccessToken)
                                                                                 Await Task.FromResult(Of Object)(Nothing)
                                                                             End Function)
            }
            Dim upn = Await GetUserIdAsync(graphClient, email)
            Dim mailbox = Await graphClient.Users(upn).Request().[Select]("MailboxSettings").GetAsync()
            Return mailbox
        End Function

        Private Async Function GetUserIdAsync(ByVal graphServiceClient As GraphServiceClient, ByVal email As String) As Task(Of String)
            Dim filter = $"startswith(userPrincipalName,'{email}')"
            Dim users = Await graphServiceClient.Users.Request().Filter(filter).GetAsync()
            Return users.CurrentPage(0).Id
        End Function
    End Class
End Namespace
