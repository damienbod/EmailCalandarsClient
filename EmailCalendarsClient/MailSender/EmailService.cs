using Microsoft.Graph;
using System.Collections.Generic;

namespace EmailCalendarsClient.MailSender
{
    static class EmailService
    {
        public static Message CreateStandardEmail(string recipient, string header,
            string body, MessageAttachmentsCollectionPage attachments)
        {
            var message = new Message
            {
                Subject = header,
                Body = new ItemBody
                {
                    ContentType = BodyType.Text,
                    Content = body
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
                },
                Attachments = attachments
            };

            return message;
        }

        public static Message CreateHtmlEmail(string recipient, string header, 
            string body, MessageAttachmentsCollectionPage attachments)
        {
            var message = new Message
            {
                Subject = header,
                Body = new ItemBody
                {
                    ContentType = BodyType.Html,
                    Content = body
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
                },
                Attachments = attachments
            };

            return message;
        }
    }
}
