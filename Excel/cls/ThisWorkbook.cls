Option Explicit

Private Sub Workbook_Open()
    Log "Workbook_Open"
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Log "Workbook_BeforeClose"
End Sub

Private Sub Log(ByVal s As String)
    Const PATH As String = "C:\temp\log.txt"
    Dim fn As Integer
    fn = FreeFile
    Open PATH For Append As #fn
    Print #fn, Join(Array(Format$(Now, "YYYY-MM-DD_hh-mm"), Environ$("USERNAME"), Environ$("COMPUTERNAME"), s), ",")
    Close fn
End Sub

Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    ' Pop up a name of a double-clicked range
    Dim ns As String
    Dim n As Name
    For Each n In ThisWorkbook.Names
        Dim v As Variant: v = Split(n.Value, "!")
        If Not (Intersect(Target, ThisWorkbook.Worksheets(Replace(v(0), "=", vbNullString)).Range(v(1))) Is Nothing) Then
            ns = ns + n.Name + vbCrLf
        End If
        ThisWorkbook.Worksheets(Replace(v(0), "=", vbNullString)).Range(v(1)).Interior.ColorIndex = 6
    Next
    MsgBox ns
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Application.screenUpdating = False
    Dim ss As Worksheets: Set ss = ThisWorkbook.Worksheets
    Dim i As Long
    For i = ss.Count To 1 Step -1
        ss(i).Activate
        ss(i).Cells(1, 1).Select
    Next
    Application.screenUpdating = True
End Sub
