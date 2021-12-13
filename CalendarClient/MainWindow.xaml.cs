using CalendarServices.CalendarClient;
using Microsoft.Identity.Client;
using Microsoft.Win32;
using System.IO;
using System.Windows;

namespace CalendarClient
{
    public partial class MainWindow : Window
    {
        AadGraphApiApplicationClient _aadGraphApiApplicationClient = new AadGraphApiApplicationClient();
        EmailService _emailService = new EmailService();

        const string SignInString = "Sign In";
        const string ClearCacheString = "Clear Cache";

        public MainWindow()
        {
            InitializeComponent();
        }

        private async void SignIn(object sender = null, RoutedEventArgs args = null)
        {
            //var accounts = await _aadGraphApiApplicationClient.GetAccountsAsync();

            //if (SignInButton.Content.ToString() == ClearCacheString)
            //{
            //    await _aadGraphApiApplicationClient.RemoveAccountsAsync();

            //    SignInButton.Content = SignInString;
            //    UserName.Content = "Not signed in";
            //    return;
            //}

            //try
            //{
            //    var account = await _aadGraphApiApplicationClient.SignIn();

            //    Dispatcher.Invoke(() =>
            //    {
            //        SignInButton.Content = ClearCacheString;
            //        SetUserName(account);
            //    });
            //}
            //catch (MsalException ex)
            //{
            //    if (ex.ErrorCode == "access_denied")
            //    {
            //        // The user canceled sign in, take no action.
            //    }
            //    else
            //    {
            //        // An unexpected error occurred.
            //        string message = ex.Message;
            //        if (ex.InnerException != null)
            //        {
            //            message += "Error Code: " + ex.ErrorCode + "Inner Exception : " + ex.InnerException.Message;
            //        }

            //        MessageBox.Show(message);
            //    }

            //    Dispatcher.Invoke(() =>
            //    {
            //        UserName.Content = "Not signed in";
            //    });
            //}
        }

        

        private async void GetCalenderEvents(object sender, RoutedEventArgs e)
        {
            var email = EmailCalendarText.Text;
            var data = await _aadGraphApiApplicationClient.GetCalanderForUser(email);
   
            var first = data.CurrentPage[0];
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
