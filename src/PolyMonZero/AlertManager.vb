Imports System.Net.Http
Imports System.Text

''' <summary>
''' Sends alert messages to a Matrix room via the Conduit homeserver.
''' Fire-and-forget async — failures are silently swallowed so alerts never crash the monitor loop.
''' </summary>
Public Class AlertManager

    Private Shared ReadOnly _Http As New HttpClient()

    Public Property MatrixHomeserver As String = ""
    Public Property MatrixToken As String = ""

    Public ReadOnly Property IsConfigured As Boolean
        Get
            Return Not String.IsNullOrWhiteSpace(MatrixHomeserver) AndAlso
                   Not String.IsNullOrWhiteSpace(MatrixToken)
        End Get
    End Property

    ''' <summary>Send a message to a Matrix room. Non-blocking — caller does not await.</summary>
    Public Sub SendAsync(ByVal roomId As String, ByVal message As String)
        If Not IsConfigured OrElse String.IsNullOrWhiteSpace(roomId) OrElse String.IsNullOrWhiteSpace(message) Then Return
        Dim homeserver = MatrixHomeserver.TrimEnd("/"c)
        Dim token = MatrixToken
        Task.Run(Async Function()
                     Try
                         Dim txnId = Guid.NewGuid().ToString("N")
                         Dim url = $"{homeserver}/_matrix/client/v3/rooms/{Uri.EscapeDataString(roomId)}/send/m.room.message/{txnId}"
                         Dim json = "{""msgtype"":""m.text"",""body"":" & JsonEscape(message) & "}"
                         Dim req As New HttpRequestMessage(HttpMethod.Put, url)
                         req.Headers.TryAddWithoutValidation("Authorization", "Bearer " & token)
                         req.Content = New StringContent(json, Encoding.UTF8, "application/json")
                         Await _Http.SendAsync(req)
                     Catch
                     End Try
                 End Function)
    End Sub

    Private Shared Function JsonEscape(s As String) As String
        If s Is Nothing Then Return """"""
        s = s.Replace("\", "\\").Replace("""", "\""").Replace(vbCr, "").Replace(vbLf, "\n").Replace(vbTab, "\t")
        Return """" & s & """"
    End Function

End Class
