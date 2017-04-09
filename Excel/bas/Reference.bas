Option Explicit

Private Sub PrintReferences()
    Dim s As String
    s = ""

    Dim r As Reference 'Microsoft Visual Basic for Applications Extensibility x.x
    For Each r In ThisWorkbook.VBProject.References
        s = s & Join(Array(r.GUID, r.Name, IIf(r.Description = "", "", r.Description), r.FullPath), vbCrLf) & String(2, vbCrLf)
    Next
    Debug.Print s
End Sub

Private Sub AddReference(ByVal dllPath As String)
    ThisWorkbook.VBProject.References.AddFromFile dllPath 'Microsoft Visual Basic for Applications Extensibility x.x
End Sub

Private Function HasReference(ByVal dllPath As String) As Boolean
    Dim r As Reference 'Microsoft Visual Basic for Applications Extensibility x.x
    For Each r In ThisWorkbook.VBProject.References
        If (r.FullPath = dllPath) Then
            ThisWorkbook.VBProject.References.Remove r
            Exit For
        End If
    Next
End Function

Private Sub RemoveReference(ByVal dllPath As String)
    Dim r As Reference 'Microsoft Visual Basic for Applications Extensibility x.x
    For Each r In ThisWorkbook.VBProject.References
        If (r.FullPath = dllPath) Then
            ThisWorkbook.VBProject.References.Remove r
            Exit For
        End If
    Next
End Sub

Private Sub AddBestAvailableDAO()
    Const PREFERRED_DLL As String = "C:\PROGRA~1\COMMON~1\MICROS~1\OFFICE12\ACEDAO.DLL"
    Const FALLBACK_DLL  As String = "C:\Program Files\Common Files\Microsoft Shared\DAO\dao360.dll"
    If (Dir(PREFERRED_DLL) <> "") And Not HasReference(PREFERRED_DLL) Then
        AddReference PREFERRED_DLL
    ElseIf Dir(FALLBACK_DLL) <> "" And Not Not HasReference(FALLBACK_DLL) Then
        RemoveReference FALLBACK_DLL
    End If
End Sub