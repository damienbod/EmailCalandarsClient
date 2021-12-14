using CalendarClient;
using EmailCalendarsClient.MailSender;
using Microsoft.Identity.Client;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Windows;

namespace GraphEmailClient
{
    public partial class MainWindow : Window
    {
        AadGraphApiDelegatedClient _aadGraphApiDelegatedClient = new AadGraphApiDelegatedClient();

        const string SignInString = "Sign In";
        const string ClearCacheString = "Clear Cache";

        public MainWindow()
        {
            InitializeComponent();
            _aadGraphApiDelegatedClient.InitClient();
        }

        private async void SignIn(object sender = null, RoutedEventArgs args = null)
        {
            var accounts = await _aadGraphApiDelegatedClient.GetAccountsAsync();

            if (SignInButton.Content.ToString() == ClearCacheString)
            {
                await _aadGraphApiDelegatedClient.RemoveAccountsAsync();

                SignInButton.Content = SignInString;
                UserName.Content = "Not signed in";
                return;
            }

            try
            {
                var account = await _aadGraphApiDelegatedClient.SignIn();

                Dispatcher.Invoke(() =>
                {
                    SignInButton.Content = ClearCacheString;
                    SetUserName(account);
                });
            }
            catch (MsalException ex)
            {
                if (ex.ErrorCode == "access_denied")
                {
                    // The user canceled sign in, take no action.
                }
                else
                {
                    // An unexpected error occurred.
                    string message = ex.Message;
                    if (ex.InnerException != null)
                    {
                        message += "Error Code: " + ex.ErrorCode + "Inner Exception : " + ex.InnerException.Message;
                    }

                    MessageBox.Show(message);
                }

                Dispatcher.Invoke(() =>
                {
                    UserName.Content = "Not signed in";
                });
            }
        }

        private async void GetCalanderForUser(object sender, RoutedEventArgs e)
        {
            var to = FilterToText.Text;
            var from = FilterFromText.Text;

            var userCalendarViewCollectionPages = await _aadGraphApiDelegatedClient
                .GetCalanderForUser(EmailRecipientText.Text, from, to);

            var allEvents = new List<FilteredEvent>();

            while (userCalendarViewCollectionPages != null && userCalendarViewCollectionPages.Count > 0)
            {
                foreach (var calenderEvent in userCalendarViewCollectionPages)
                {
                    var filteredEvent = new FilteredEvent
                    {
                        ShowAs = calenderEvent.ShowAs,
                        Sensitivity = calenderEvent.Sensitivity,
                        Start = calenderEvent.Start,
                        End = calenderEvent.End,
                        Subject = calenderEvent.Subject,
                        IsAllDay = calenderEvent.IsAllDay,
                        Location = calenderEvent.Location
                    };
                    allEvents.Add(filteredEvent);
                }
                //check if next result page is available
                if (userCalendarViewCollectionPages.NextPageRequest == null)
                    break;
            }

            CalendarDataText.Text = JsonSerializer.Serialize(allEvents);
        }


        private void SetUserName(IAccount userInfo)
        {
            string userName = null;

            if (userInfo != null)
            {
                userName = userInfo.Username;
            }

            if (userName == null)
            {
                userName = "Not identified";
            }

            UserName.Content = userName;
        }
    }
}
