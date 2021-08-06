using EmailCalendarsClient.MailSender;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.Win32;
using System;
using System.IO;
using System.Net.Http;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows;

namespace GraphEmailClient
{
    public partial class MainWindow : Window
    {
        AadNativeClient _aadNativeClient = new AadNativeClient();
        MessageAttachmentsCollectionPage _messageAttachmentsCollectionPage = new MessageAttachmentsCollectionPage();
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

        private static async Task DisplayErrorMessage(HttpResponseMessage httpResponse)
        {
            string failureDescription = await httpResponse.Content.ReadAsStringAsync();
            if (failureDescription.StartsWith("<!DOCTYPE html>"))
            {
                string path = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
                string errorFilePath = Path.Combine(path, "error.html");
                System.IO.File.WriteAllText(errorFilePath, failureDescription);
                System.Diagnostics.Process.Start(errorFilePath);
            }
            else
            {
                MessageBox.Show($"{httpResponse.ReasonPhrase}\n {failureDescription}", "An error occurred while sending an email", MessageBoxButton.OK);
            }
        }

        private async void SendEmail(object sender, RoutedEventArgs e)
        {
            var message = EmailService.CreateStandardEmail(EmailRecipientText.Text, 
                EmailHeader.Text, EmailBody.Text, _messageAttachmentsCollectionPage);

            await _aadNativeClient.SendEMailAsync(message);

            _messageAttachmentsCollectionPage.Clear();
        }

        private async void SendHtmlEmail(object sender, RoutedEventArgs e)
        {
            var messageHtml = EmailService.CreateHtmlEmail(EmailRecipientText.Text,
                EmailHeader.Text, EmailBody.Text, _messageAttachmentsCollectionPage);

            await _aadNativeClient.SendEMailAsync(messageHtml);

            _messageAttachmentsCollectionPage.Clear();
        }

        private void AddAttachment(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog();
            string fileAsString = string.Empty;
            if (dlg.ShowDialog() == true)
            {
                new FileInfo(dlg.FileName);
                using (Stream s = dlg.OpenFile())
                {
                    TextReader reader = new StreamReader(s);
                    fileAsString = reader.ReadToEnd();
                }

                _messageAttachmentsCollectionPage.Add(new FileAttachment
                { 
                    Name = dlg.FileName,
                    ContentBytes = EncodeTobase64Bytes(fileAsString)
                });

            }
        }

        static public byte[] EncodeTobase64Bytes(string toEncode)
        {
            byte[] toEncodeAsBytes = System.Text.ASCIIEncoding.ASCII.GetBytes(toEncode);
            string base64String = System.Convert.ToBase64String(toEncodeAsBytes);
            var returnValue = Convert.FromBase64String(base64String);
            return returnValue;
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
