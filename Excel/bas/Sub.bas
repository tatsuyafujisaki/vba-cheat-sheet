Option Explicit

Private Sub ToDate(r As Range)
    Dim SingleCell As Range
    For Each SingleCell In r
        If IsDate(SingleCell.Value) Then SingleCell.Value = CDate(SingleCell.Value)
    Next
End Sub

Private Sub MkDirIfNotExist(ByVal path As String)
    If Dir(path, vbDirectory) = "" Then MkDir path
End Sub

Private Sub Format()
    With Me
        With .PageSetup
            .PrintArea = Me.UsedRange.address
            .Orientation = IIf(Me.UsedRange.columns.Count < Me.UsedRange.Rows.Count, xlPortrait, xlLandscape)
            .PaperSize = xlPaperA4
            .Zoom = False 'Needs to be False to enable FitToPagesTall and FitToPagesWide
            .FitToPagesTall = False
            .FitToPagesWide = 1
            .HeaderMargin = 0
            .FooterMargin = 0
            .TopMargin = 0
            .BottomMargin = 0
            .LeftMargin = 0
            .RightMargin = 0
            .PrintGridlines = True
        End With

        With .Cells.Font
            .name = "Segoe UI"
            .Size = 10
        '    .Rows.AutoFit
        '    .Columns.AutoFit
        End With
    End With
End Sub

Private Sub ExampleOfGetFilename(ByVal filter As String)
    'e.g. filter = "Templates (*.xlt), *.xlt"
    Dim fileName As String: fileName = Application.GetOpenFilename(filter)
    If fileName = "False" Then
        MsgBox "Please select data file"
        Exit Sub
    End If
    Dim wb As Workbook: Set wb = Workbooks.Open(fileName, ReadOnly:=True)
    Dim ws As Worksheet: Set ws = wb.Sheets(EQ_SHEET_NAME)
End Sub

Private Sub CloseWithoutPrompt(wb As Workbook)
    Application.DisplayAlerts = False
    wb.Close False
    Application.DisplayAlerts = True
End Sub

Private Sub SaveCloseWithoutPrompt(wb As Workbook, ByVal fileName As String)
    Application.DisplayAlerts = False
    wb.Close True, fileName
    Application.DisplayAlerts = True
End Sub

Private Sub SaveWithoutPrompt(wb As Workbook, ByVal fileName As String)
    Application.DisplayAlerts = False
    wb.SaveAs fileName
    Application.DisplayAlerts = True
End Sub

'Usage: ExportVBComponents Array("Excel1.xlsx", "Module1.bas", "ClassModule1.cls", "Form1.frm")
Private Sub ExportVBComponents(files As Variant)
    Const DESTINATION_WORKBOOK_NAME As String = "production.xls"
    Dim fso As New FileSystemObject
    Dim file As Variant
    With Workbooks.Open(ThisWorkbook.path & "\" & DESTINATION_WORKBOOK_NAME)
        For Each file In files
            Dim path As String: path = Environ("temp") & "\" & file
            Dim module As String: module = fso.GetBaseName(file)
            ThisWorkbook.VBProject.VBComponents(module).Export path
            On Error Resume Next
            Dim vbc As VBComponent: Set vbc = .VBProject.VBComponents(module) 'Microsoft Visual Basic for Applications Extensibility x.x
            If Err.Number = 0 Then .VBProject.VBComponents.Remove vbc
            On Error GoTo 0
            .VBProject.VBComponents.Import path
            Kill path
        Next
        .Close True
    End With
End Sub

Private Sub ImportSheets(ByVal before As String, ByVal after As String, ParamArray wsNames() As Variant)
    Const SOURCE_WORKBOOK_NAME As String = "release.xls"
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets(wsNames).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    With Workbooks.Open(ThisWorkbook.path & "\" & SOURCE_WORKBOOK_NAME, ReadOnly:=True)
        If before <> "" Then .Sheets(wsNames).Copy before:=ThisWorkbook.Sheets(before)
        If after <> "" Then .Sheets(wsNames).Copy after:=ThisWorkbook.Sheets(after)
        .Close False
    End With
End Sub
