Imports Microsoft.Graph
Imports System
Imports System.Collections.Generic
Imports System.IO

Namespace EmailCalendarsClient.MailSender
    Class EmailService
        Private MessageAttachmentsCollectionPage As MessageAttachmentsCollectionPage = New MessageAttachmentsCollectionPage()

        Public Function CreateStandardEmail(ByVal recipient As String, ByVal header As String, ByVal body As String) As Message
            Dim message = New Message With {
                .Subject = header,
                .Body = New ItemBody With {
                    .ContentType = BodyType.Text,
                    .Content = body
                },
                .ToRecipients = New List(Of Recipient)() From {
                    New Recipient With {
                        .EmailAddress = New EmailAddress With {
                            .Address = recipient
                        }
                    }
                },
                .Attachments = MessageAttachmentsCollectionPage
            }
            Return message
        End Function

        Public Function CreateHtmlEmail(ByVal recipient As String, ByVal header As String, ByVal body As String) As Message
            Dim message = New Message With {
                .Subject = header,
                .Body = New ItemBody With {
                    .ContentType = BodyType.Html,
                    .Content = body
                },
                .ToRecipients = New List(Of Recipient)() From {
                    New Recipient With {
                        .EmailAddress = New EmailAddress With {
                            .Address = recipient
                        }
                    }
                },
                .Attachments = MessageAttachmentsCollectionPage
            }
            Return message
        End Function

        Public Sub AddAttachment(ByVal rawData As Byte(), ByVal filePath As String)
            MessageAttachmentsCollectionPage.Add(New FileAttachment With {
                .Name = Path.GetFileName(filePath),
                .ContentBytes = EncodeTobase64Bytes(rawData)
            })
        End Sub

        Public Sub ClearAttachments()
            MessageAttachmentsCollectionPage.Clear()
        End Sub

        Public Shared Function EncodeTobase64Bytes(ByVal rawData As Byte()) As Byte()
            Dim base64String As String = System.Convert.ToBase64String(rawData)
            Dim returnValue = Convert.FromBase64String(base64String)
            Return returnValue
        End Function
    End Class
End Namespace
