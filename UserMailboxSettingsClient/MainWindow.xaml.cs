using CalendarServices.UserMailboxSettingsClient;
using System.Collections.Generic;
using System.Text.Json;
using System.Windows;

namespace UserMailboxSettingsClient
{
    public partial class MainWindow : Window
    {
        AadGraphApiApplicationClient _aadGraphApiApplicationClient = new AadGraphApiApplicationClient();

        public MainWindow()
        {
            InitializeComponent();
        }

        private async void GetUserMailboxSettings(object sender, RoutedEventArgs e)
        {
            var user = await _aadGraphApiApplicationClient
                .GetUserMailboxSettings(EmailText.Text);

            UserMailboxSettingsText.Text = JsonSerializer.Serialize(user.MailboxSettings);
        }
    }
}
