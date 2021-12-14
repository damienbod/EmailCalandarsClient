Imports Microsoft.Identity.Client
Imports System.Text.Json
Imports VBUserMailboxSettingsClient.MailboxSettings

Partial Public Class MainWindow
    Inherits Window

    Private _aadGraphApiDelegatedClient As AadGraphApiDelegatedClient = New AadGraphApiDelegatedClient()
    Const SignInString As String = "Sign In"
    Const ClearCacheString As String = "Clear Cache"

    Public Sub New()
        InitializeComponent()
        _aadGraphApiDelegatedClient.InitClient()
    End Sub

    Private Async Sub SignIn(ByVal Optional sender As Object = Nothing, ByVal Optional args As RoutedEventArgs = Nothing)
        Dim accounts = Await _aadGraphApiDelegatedClient.GetAccountsAsync()

        If SignInButton.Content.ToString() = ClearCacheString Then
            Await _aadGraphApiDelegatedClient.RemoveAccountsAsync()
            SignInButton.Content = SignInString
            UserName.Content = "Not signed in"
            Return
        End If

        Try
            Dim account = Await _aadGraphApiDelegatedClient.SignIn()
            Dispatcher.Invoke(Function()
                                  SignInButton.Content = ClearCacheString
                                  SetUserName(account)
                              End Function)
        Catch ex As MsalException

            If ex.ErrorCode = "access_denied" Then
            Else
                Dim message As String = ex.Message

                If ex.InnerException IsNot Nothing Then
                    message += "Error Code: " & ex.ErrorCode & "Inner Exception : " + ex.InnerException.Message
                End If

                MessageBox.Show(message)
            End If

            Dispatcher.Invoke(Function()
                                  UserName.Content = "Not signed in"
                              End Function)
        End Try
    End Sub

    Private Async Sub GetMailBoxSettingsforEmail(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Dim user = Await _aadGraphApiDelegatedClient.GetUserMailboxSettings(EmailRecipientText.Text)
        EmailBody.Text = JsonSerializer.Serialize(user.MailboxSettings)
    End Sub

    Private Sub SetUserName(ByVal userInfo As IAccount)
        Dim userNameString As String = Nothing

        If userInfo IsNot Nothing Then
            userNameString = userInfo.Username
        End If

        If userNameString Is Nothing Then
            userNameString = "Not identified"
        End If

        UserName.Content = userNameString
    End Sub
End Class