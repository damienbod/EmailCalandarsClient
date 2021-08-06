using Microsoft.Graph;
using System.Collections.Generic;

namespace EmailCalendarsClient.MailSender
{
    static class EmailService
    {
        public static Message CreateStandardEmail(string recipient)
        {
            var message = new Message
            {
                Subject = "Meet for lunch?",
                Body = new ItemBody
                {
                    ContentType = BodyType.Text,
                    Content = "The new cafeteria is open."
                },
                ToRecipients = new List<Recipient>()
                {
                    new Recipient
                    {
                        EmailAddress = new EmailAddress
                        {
                            Address = recipient
                        }
                    }
                }
            };

            return message;
        }

        public static Message CreateHtmlEmail(string recipient)
        {
            var message = new Message
            {
                Subject = "Meet for lunch?",
                Body = new ItemBody
                {
                    ContentType = BodyType.Html,
                    Content = "<p>The new <b>cafeteria</b> is open.</p>"
                },
                ToRecipients = new List<Recipient>()
                {
                    new Recipient
                    {
                        EmailAddress = new EmailAddress
                        {
                            Address = recipient
                        }
                    }
                }
            };

            return message;
        }
    }
}
