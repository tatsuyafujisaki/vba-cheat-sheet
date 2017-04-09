Option Explicit

Private Sub SendPlainEmail(ByVal from As String, ByVal to1 As String, ByVal cc As String, ByVal bcc As String, ByVal subject As String, ByVal body As String)
    With New CDO.message 'Microsoft CDO for Windows 2000 Library
        With .Configuration.Fields
            .item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPort
            .item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "localhost"
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

Private Sub CreatePlainEmail(ByVal to1 As String, ByVal subject As String, ByVal body As String, Optional atts As Collection = Nothing)
    With New Outlook.Application 'Microsoft Outlook x.x Object Library
        With .CreateItem(0)
            .To = to1
            .subject = subject
            .body = body
            If Not (atts Is Nothing) Then
                For Each att In atts
                    .attachments.Add att
                Next
            End If
            .Display
        End With
    End With
End Sub

Private Sub CreateHtmlEmail(ByVal to1 As String, ByVal subject As String, ByVal body As String, Optional atts As Collection = Nothing)
    Const HEAD As String = "<head><style>p{font:9pt ""Meiryo UI"";}</style></head>"
    With New Outlook.Application 'Microsoft Outlook x.x Object Library
        With .CreateItem(0)
            .To = to1
            .subject = subject
            .BodyFormat = olFormatHTML
            .HTMLBody = HEAD & body 'body is like "<p>Hello world!</p>"
            If Not (atts Is Nothing) Then
                For Each att In atts
                    .attachments.Add att
                Next
            End If
            .Display
        End With
    End With
End Sub