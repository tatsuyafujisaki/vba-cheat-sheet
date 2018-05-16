Option Explicit

Sub PutInClipboard(ByVal s As String)
    With New DataObject 'Microsoft Forms 2.0 Object Library (or manually reference C:\WINDOWS\system32\FM20.DLL)
        .SetText s
        .PutInClipboard
    End With
End Sub

Function GetFromClipboard() As String
    With New DataObject 'Microsoft Forms 2.0 Object Library (or manually reference C:\WINDOWS\system32\FM20.DLL)
        .GetFromClipboard
        GetFromClipboard = .GetText
    End With
End Function
