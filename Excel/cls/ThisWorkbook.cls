Option Explicit

Private Sub Workbook_Open()
    Log "Workbook_Open"
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Log "Workbook_BeforeClose"
End Sub

Private Sub Log(ByVal s As String)
    Const Path As String = "C:\Foo.log"

    Dim fileNumber As Integer
    fileNumber = FreeFile

    Open Path For Append As #fileNumber
    Print #fileNumber, Join$(Array(Format$(Now, "yyyy-mm-dd hh:mm"), Environ$("ComputerName"), Application.UserName, s), "|")
    Close fileNumber
End Sub

' Show the name of a double-clicked range
Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    Dim names As String
    Dim name1 As Name

    For Each name1 In Me.names
        Dim v As Variant
        v = Split(name1.Value, "!")
        
        If Not (Intersect(Target, Me.Worksheets(Replace$(v(0), "=", vbNullString)).Range(v(1))) Is Nothing) Then
            names = names + name1.Name + vbNewLine
        End If
        Me.Worksheets(Replace$(v(0), "=", vbNullString)).Range(v(1)).Interior.ColorIndex = 6
    Next
    MsgBox names
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Dim o As Optimizer
    Set o = New Optimizer

    With Me
        Dim i As Long
        For i = Me.Worksheets.Count To 1 Step -1
            Me.Worksheets(i).Activate
            Me.Worksheets(i).Cells(1, 1).Select
        Next
    End With
End Sub
