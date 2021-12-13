using CalendarServices.CalendarClient;
using System.Collections.Generic;
using System.Text.Json;
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

            var userCalendarViewCollectionPages = await _aadGraphApiApplicationClient
                .GetCalanderForUser(EmailCalendarText.Text, from, to);

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
    }
}
