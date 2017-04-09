Option Explicit

' If "Microsoft Forms x.x Object Library" is not in the References list, manually reference C:\WINDOWS\system32\FM20.DLL

Private Sub PutInClipboard(ByVal s As String)
    With New DataObject 'Microsoft Forms x.x Object Library
        .SetText s
        .PutInClipboard
    End With
End Sub

Private Function GetFromClipboard() As String
    With New DataObject 'Microsoft Forms x.x Object Library
        .GetFromClipboard
        GetFromClipboard = .GetText
    End With
End Function
