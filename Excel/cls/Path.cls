Option Explicit

Public Value As String

Public Function Join(ByVal dir As String) As path
    With New FileSystemObject
        Value = .BuildPath(Value, dir)
        Set Join = Me
    End With
End Function

' Usasge
' Dim p As Path
' Set p = New Path
'
' The following prints C:\Foo\Bar
' Debug.Print p.Join("C:").Join("Foo").Join("Bar").Value
