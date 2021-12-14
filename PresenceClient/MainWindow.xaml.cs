using Microsoft.Identity.Client;
using Microsoft.Win32;
using System;
using System.IO;
using System.Windows;
using PresenceClient.Presence;
using System.Text.Json;
using System.Collections.Generic;
using Microsoft.Graph;

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

        private async void GetPresenceforEmail(object sender, RoutedEventArgs e)
        {
            var cloudCommunicationPages = await _aadGraphApiDelegatedClient.GetPresenceAsync(EmailRecipientText.Text);

            var allPresenceItems = new List<Presence>();

            while (cloudCommunicationPages != null && cloudCommunicationPages.Count > 0)
            {
                foreach (var presence in cloudCommunicationPages)
                {
                    allPresenceItems.Add(presence);
                }
                //check if next result page is available
                if (cloudCommunicationPages.NextPageRequest == null)
                    break;
            }

            EmailBody.Text = JsonSerializer.Serialize(allPresenceItems);
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
