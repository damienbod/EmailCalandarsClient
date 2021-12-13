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
            var to = FilterToText.Text;
            var from = FilterFromText.Text;

            var data = await _aadGraphApiApplicationClient
                .GetCalanderForUser(EmailCalendarText.Text, to, from);
   
            var first = data.CurrentPage[0];
        }
    }
}
