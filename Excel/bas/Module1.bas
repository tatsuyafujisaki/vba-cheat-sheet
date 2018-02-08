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

Private Sub ExampleOfGetFilename(ByVal filter As String)
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

'Usage: ExportVBComponents Array("Excel1.xlsx", "Module1.bas", "ClassModule1.cls", "Form1.frm")
Private Sub ExportVBComponents(ByVal files As Variant)
    Const DestinationWorkbookName As String = "Production.xlsx"
    Dim fso As New FileSystemObject
    Dim file As Variant
    With Workbooks.Open(ThisWorkbook.Path & Application.PathSeparator & DestinationWorkbookName)
        For Each file In files
            Dim path As String
            path = Environ$("temp") & Application.PathSeparator & file

            Dim module As String
            module = fso.GetBaseName(file)

            ThisWorkbook.VBProject.VBComponents(module).Export path

            On Error Resume Next

            Dim vbc As VBComponent 'Microsoft Visual Basic for Applications Extensibility x.x
            Set vbc = .VBProject.VBComponents(module)

            If Err.Number = 0 Then
                .VBProject.VBComponents.Remove vbc
            End If

            On Error GoTo 0

            .VBProject.VBComponents.Import path
            Kill path
        Next
        .Close True
    End With
End Sub

Private Sub ImportSheets(ByVal before As String, ByVal after As String, ParamArray wsNames() As Variant)
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