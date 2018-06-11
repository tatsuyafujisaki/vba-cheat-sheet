Option Explicit

Sub Wait(ByVal time As String)
   Application.Wait Now + TimeValue(time)
End Sub

Sub ToDate(ByVal r As Range)
    Dim SingleCell As Range
    For Each SingleCell In r
        If IsDate(SingleCell.Value) Then
            SingleCell.Value = CDate(SingleCell.Value)
        End If
    Next
End Sub

Sub ExampleOfGetFilename(ByVal filter As String)
    'e.g. filter = "Templates (*.xlt), *.xlt"
    Dim fileName As String
    fileName = Application.GetOpenFilename(filter)

    If fileName = "False" Then
        MsgBox "Please select data file"
        Exit Sub
    End If

    Dim wb As Workbook
    Set wb = Workbooks.Open(fileName, ReadOnly:=True)

    Dim ws As Worksheet
    Set ws = wb.Worksheets(fileName)
End Sub

Sub ImportSheets(ByVal before As String, ByVal after As String, ParamArray wsNames() As Variant)
    Const SourceWorkbookName As String = "Release.xlsx"
    Application.DisplayAlerts = False

    On Error Resume Next
    ThisWorkbook.Worksheets(wsNames).Delete
    On Error GoTo 0

    Application.DisplayAlerts = True
    With Workbooks.Open(ThisWorkbook.path & Application.PathSeparator & SourceWorkbookName, ReadOnly:=True)
        If before <> vbNullString Then
            .Worksheets(wsNames).Copy before:=ThisWorkbook.Worksheets(before)
        End If

        If after <> vbNullString Then
            .Worksheets(wsNames).Copy after:=ThisWorkbook.Worksheets(after)
        End If

        .Close False
    End With
End Sub