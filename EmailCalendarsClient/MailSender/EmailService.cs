using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.IO;

namespace EmailCalendarsClient.MailSender
{
    class EmailService
    {
        MessageAttachmentsCollectionPage MessageAttachmentsCollectionPage = new MessageAttachmentsCollectionPage();

        public Message CreateStandardEmail(string recipient, string header, string body)
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
                Attachments = MessageAttachmentsCollectionPage
            };

            return message;
        }

        public Message CreateHtmlEmail(string recipient, string header, string body)
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
                Attachments = MessageAttachmentsCollectionPage
            };

            return message;
        }

        public void AddAttachment(string fileAsString, string filePath)
        {
            MessageAttachmentsCollectionPage.Add(new FileAttachment
            {
                Name = Path.GetFileName(filePath),
                ContentBytes = EncodeTobase64Bytes(fileAsString)
            });
        }

        public void ClearAttachments()
        {
            MessageAttachmentsCollectionPage.Clear();
        }

        static public byte[] EncodeTobase64Bytes(string toEncode)
        {
            byte[] toEncodeAsBytes = System.Text.ASCIIEncoding.ASCII.GetBytes(toEncode);
            string base64String = System.Convert.ToBase64String(toEncodeAsBytes);
            var returnValue = Convert.FromBase64String(base64String);
            return returnValue;
        }
    }
}
