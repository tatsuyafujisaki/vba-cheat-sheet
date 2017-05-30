Option Explicit

Private Sub SendPlainEmail(ByVal from As String, ByVal to1 As String, ByVal cc As String, ByVal bcc As String, ByVal subject As String, ByVal body As String)
    With New CDO.Message 'Microsoft CDO for Windows 2000 Library
        With .Configuration.Fields
            .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPort
            .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "localhost"
            .Update
        End With
        .from = from
        .To = to1
        .cc = cc
        .bcc = bcc
        .subject = subject
        .TextBody = body
        .send
    End With
End Sub

Private Sub CreatePlainEmail(ByVal to1 As String, ByVal subject As String, ByVal body As String, Optional ByVal attachments As Collection = Nothing)
    Dim attachment As Object

    With New Outlook.Application 'Microsoft Outlook x.x Object Library
        With .CreateItem(0)
            .To = to1
            .subject = subject
            .body = body
            If Not (attachments Is Nothing) Then
                For Each attachment In attachments
                    .attachments.Add attachment
                Next
            End If
            .Display
        End With
    End With
End Sub

Private Sub CreateHtmlEmail(ByVal to1 As String, ByVal subject As String, ByVal body As String, Optional ByVal attachments As Collection = Nothing)
    Const Head As String = "<head><style>p{font:9pt ""Meiryo UI"";}</style></head>"
    
    Dim attachment As Object
    
    With New Outlook.Application 'Microsoft Outlook x.x Object Library
        With .CreateItem(0)
            .To = to1
            .subject = subject
            .BodyFormat = olFormatHTML
            .HTMLBody = Head & body 'body is like "<p>Hello world!</p>"
            If Not (attachments Is Nothing) Then
                For Each attachment In attachments
                    .attachments.Add attachment
                Next
            End If
            .Display
        End With
    End With
End Sub
