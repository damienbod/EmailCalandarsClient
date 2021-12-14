Imports Microsoft.Graph

Public Class FilteredEvent
    Public Property Start As DateTimeTimeZone
    Public Property [End] As DateTimeTimeZone
    Public Property Subject As String
    Public Property Location As Location
    Public Property Sensitivity As Sensitivity?
    Public Property ShowAs As FreeBusyStatus?
    Public Property IsAllDay As Boolean?
End Class
