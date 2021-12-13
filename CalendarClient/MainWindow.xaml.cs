using CalendarServices.CalendarClient;
using System.Windows;

namespace CalendarClient
{
    public partial class MainWindow : Window
    {
        AadGraphApiApplicationClient _aadGraphApiApplicationClient = new AadGraphApiApplicationClient();

        public MainWindow()
        {
            InitializeComponent();
        }

        private async void GetCalenderEvents(object sender, RoutedEventArgs e)
        {
            var email = EmailCalendarText.Text;
            var data = await _aadGraphApiApplicationClient.GetCalanderForUser(email);
   
            var first = data.CurrentPage[0];
        }
    }
}
