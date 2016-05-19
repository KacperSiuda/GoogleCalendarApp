using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using System.Threading;
using Google.Apis.Util.Store;
using Google.Apis.Calendar.v3;
using Google.Apis.Services;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Google.Apis.Calendar.v3.Data;

namespace GoogleCalendarApiTest_v1.GoogleAPI
{
    public class GoogleApiConector
    {
        private static string[] Scopes = { CalendarService.Scope.Calendar };
        private static string ApplicationName = "Google Calendar API .NET Quickstart";
        private UserCredential credential;

        public void CreateCredential()
        {
            using (var stream = new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/calendar.json");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;

            }
        }
        public void CreateEvent(string summary, DateTime start, DateTime end)
        {
            var service = new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
            Event newEvent = new Event()
            {
                Summary = summary,
                Start = new EventDateTime()
                {
                    DateTime = DateTime.Parse(start.ToString())
                },
                End = new EventDateTime()
                {
                    DateTime = DateTime.Parse(end.ToString())
                }
                

            };
            String calendarId = "primary";
            EventsResource.InsertRequest request = service.Events.Insert(newEvent, calendarId);
            Event createdEvent = request.Execute();

        }

        public Dictionary<string, Event> GetEventList()
        {
            Dictionary<string, Event> eventDictionary = new Dictionary<string, Event>();
            var service = new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName
            });
            EventsResource.ListRequest request = service.Events.List("primary");
            request.TimeMin = DateTime.Now;
            request.ShowDeleted = false;
            request.SingleEvents = true;
            request.MaxResults = 100;
            request.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;

            Events events = request.Execute();

            if (events.Items != null && events.Items.Count > 0)
            {
                foreach (var eventItem in events.Items)
                {
                    string when = eventItem.Start.DateTime.ToString();
                    if (String.IsNullOrEmpty(when))
                    {
                        when = eventItem.Start.Date;
                    }
                    var event_date = eventItem.Summary + " " + when;
                    eventDictionary.Add(eventItem.Id, eventItem);
                }
            }
            return eventDictionary;

        }
        public void DeleteEvent(string eventId)
        {
            var service = new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName
            });
            service.Events.Delete("primary", eventId).Execute();
        }

    }
}
