using EmailCalendarsClient.MailSender;
using Microsoft.Identity.Client;
using Microsoft.Win32;
using System;
using System.IO;
using System.Windows;

namespace GraphEmailClient
{
    public partial class MainWindow : Window
    {
        AadNativeClient _aadNativeClient = new AadNativeClient();
        EmailService _emailService = new EmailService();
        const string SignInString = "Sign In";
        const string ClearCacheString = "Clear Cache";

        public MainWindow()
        {
            InitializeComponent();
            _aadNativeClient.InitClient();
        }

        private async void SignIn(object sender = null, RoutedEventArgs args = null)
        {
            var accounts = await _aadNativeClient.GetAccountsAsync();

            if (SignInButton.Content.ToString() == ClearCacheString)
            {
                await _aadNativeClient.RemoveAccountsAsync();

                SignInButton.Content = SignInString;
                UserName.Content = "Not signed in";
                return;
            }

            try
            {
                var account = await _aadNativeClient.SignIn();

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

        private async void SendEmail(object sender, RoutedEventArgs e)
        {
            var message = _emailService.CreateStandardEmail(EmailRecipientText.Text, 
                EmailHeader.Text, EmailBody.Text);

            await _aadNativeClient.SendEMailAsync(message);
            _emailService.ClearAttachments();
        }

        private async void SendHtmlEmail(object sender, RoutedEventArgs e)
        {
            var messageHtml = _emailService.CreateHtmlEmail(EmailRecipientText.Text,
                EmailHeader.Text, EmailBody.Text);

            await _aadNativeClient.SendEMailAsync(messageHtml);
            _emailService.ClearAttachments();
        }

        private void AddAttachment(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog();
            if (dlg.ShowDialog() == true)
            {
                byte[] data = File.ReadAllBytes(dlg.FileName);
                _emailService.AddAttachment(data, dlg.FileName);
            }
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
