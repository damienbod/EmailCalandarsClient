Imports Microsoft.Identity.Client
Imports System.Text.Json
Imports VBCalendarClient.Calendars

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

    Private Async Sub GetCalanderForUser(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Dim [to] = FilterToText.Text
        Dim from = FilterFromText.Text
        Dim userCalendarViewCollectionPages = Await _aadGraphApiDelegatedClient.GetCalanderForUser(EmailRecipientText.Text, from, [to])
        Dim allEvents = New List(Of FilteredEvent)()

        While userCalendarViewCollectionPages IsNot Nothing AndAlso userCalendarViewCollectionPages.Count > 0

            For Each calenderEvent In userCalendarViewCollectionPages
                Dim filteredEvent = New FilteredEvent With {
                    .ShowAs = calenderEvent.ShowAs,
                    .Sensitivity = calenderEvent.Sensitivity,
                    .Start = calenderEvent.Start,
                    .[End] = calenderEvent.[End],
                    .Subject = calenderEvent.Subject,
                    .IsAllDay = calenderEvent.IsAllDay,
                    .Location = calenderEvent.Location
                }
                allEvents.Add(filteredEvent)
            Next

            If userCalendarViewCollectionPages.NextPageRequest Is Nothing Then Exit While
        End While

        CalendarDataText.Text = JsonSerializer.Serialize(allEvents)
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

