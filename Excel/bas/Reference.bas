Option Explicit

Sub PrintReferences()
    Dim s As String
    s = vbNullString

    Dim r As Reference 'Microsoft Visual Basic for Applications Extensibility 5.3
    For Each r In ThisWorkbook.VBProject.References
        s = s & Join(Array(r.GUID, r.Name, IIf(r.Description = vbNullString, vbNullString, r.Description), r.FullPath), vbCrLf) & String(2, vbCrLf)
    Next
    Debug.Print s
End Sub

Sub AddReference(ByVal dllPath As String)
    ThisWorkbook.VBProject.References.AddFromFile dllPath 'Microsoft Visual Basic for Applications Extensibility x.x
End Sub

Function HasReference(ByVal dllPath As String) As Boolean
    Dim r As Reference 'Microsoft Visual Basic for Applications Extensibility x.x
    For Each r In ThisWorkbook.VBProject.References
        If (r.FullPath = dllPath) Then
            ThisWorkbook.VBProject.References.Remove r
            Exit For
        End If
    Next
End Function

Sub RemoveReference(ByVal dllPath As String)
    Dim r As Reference 'Microsoft Visual Basic for Applications Extensibility x.x
    For Each r In ThisWorkbook.VBProject.References
        If (r.FullPath = dllPath) Then
            ThisWorkbook.VBProject.References.Remove r
            Exit For
        End If
    Next
End Sub

Sub AddBestAvailableDAO()
    Const PREFERRED_DLL As String = "C:\PROGRA~1\COMMON~1\MICROS~1\OFFICE12\ACEDAO.DLL"
    Const FALLBACK_DLL  As String = "C:\Program Files\Common Files\Microsoft Shared\DAO\dao360.dll"
    If (dir(PREFERRED_DLL) <> vbNullString) And Not HasReference(PREFERRED_DLL) Then
        AddReference PREFERRED_DLL
    ElseIf dir(FALLBACK_DLL) <> vbNullString And Not Not HasReference(FALLBACK_DLL) Then
        RemoveReference FALLBACK_DLL
    End If
End Sub
