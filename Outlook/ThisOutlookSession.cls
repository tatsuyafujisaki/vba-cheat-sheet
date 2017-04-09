'Detector of receiving an email in Inbox
Private Sub Application_NewMail()
    Dim o As Object
    ' Use olFolderDeletedItems instead if "Deleted Items" is your interest.
    For Each o In GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).Items
        'Skip MeetingItem, etc
        If TypeName(o) = "MailItem" Then
            Dim mi As MailItem
            Set mi = o
            If mi.UnRead And InStr(mi.Subject, "This is a subject.") Then
                mi.Unread = False
                With New FileSystemObject
                    .CreateTextFile "C:\temp\_next"
                End With
            End If
        End If
    Next
End Sub

'Email forwarder
Private Sub Application_NewMail()
    Dim mi As MailItem
    For Each mi In GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).Folders("SPAM").Items
        If mi.UnRead Then
            mi.UnRead = False
            With mi.Forward
                .To = "foo@example.com"
                .CC = "bar@example.com"
                .Subject = mi.Subject
                .body = mi.body
                .Send
            End With
        End If
    Next
End Sub