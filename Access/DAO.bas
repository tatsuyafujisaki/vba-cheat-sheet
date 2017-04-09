Option Explicit

Private Sub ExcelSample()
    Dim ew As New ExcelWrapper
    With ew.excel
        .SheetsInNewWorkbook = 1
        .ScreenUpdating = False
        .DisplayAlerts = False
        Dim wb As Workbook
        Set wb = .Workbooks.Add
        Dim ws As Worksheet
        Set ws = wb.Sheets(1)
        .Visible = True
        .DisplayAlerts = True
        .ScreenUpdating = True
    End With
End Sub

Private Sub Form_Open(Cancel As Integer)
    DoCmd.RunCommand acCmdSizeToFitForm
End Sub

Private Sub Form_Timer()
    Const START_UNAVAILABLE As Long = 0
    Const END_UNAVAILABLE As Long = 4
    Dim STOPPER As String
    STOPPER = CurrentProject.path & "\stop.txt"
    If (START_UNAVAILABLE <= Hour(Time) And Hour(Time) <= END_UNAVAILABLE) Or (Dir(STOPPER) <> "") Then Application.Quit acQuitSaveNone
End Sub

Private Function IsProduction() As Boolean
    Const PRD_DIR As String = "path/to/production"
    IsProduction = CurrentProject.Path = PRD_DIR
End Function

Private Sub LoosenTable(ByVal tableName As String)
    DoCmd.Close acTable, tableName
    Dim field As field
    With CurrentDb 'instantiate db
        For Each field In .TableDefs(tableName).fields
            With field
                .Required = False
                On Error Resume Next
                .AllowZeroLength = True
                .ValidationRule = Empty
                On Error GoTo 0
            End With
        Next
    End With
End Sub

Private Function GetTable(ByVal table As String) As Variant
    With CurrentDb.OpenRecordset(table)
        GetTable = WorksheetFunction.Transpose(.GetRows(.RecordCount))
        .Close
    End With
End Function

Private Function SelectQuery(ByVal queryName As String) As Variant
    With CurrentDb.QueryDefs(queryName).OpenRecordset
        If .EOF Then
            MsgBox "No record"
            End
        End If
        .MoveLast
        Dim rc As Long
        rc = .RecordCount
        .MoveFirst
        SelectQuery = .GetRows(rc)
        .Close
    End With
End Function

Private Sub SelectQueryOneRow(ByVal queryName As String, result As Dictionary)
    With CurrentDb.QueryDefs(queryName).OpenRecordset
        If .EOF Then
            MsgBox "No record"
            End
        End If
        Dim key As Variant
        For Each key In result.Keys
            result(key) = .Fields(key)
        Next
        .Close
    End With
End Sub

Private Sub SelectQueryToRange(ByVal queryName As String, r As Range, ByVal includeHeader As Boolean)
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.QueryDefs(queryName).OpenRecordset
    If rs.EOF Then
        MsgBox "No record found"
        Err.Raise 63 'Bad record number (https://support.microsoft.com/kb/146864)
        'This error is to let the parent do Excel.Quit
    End If
    If includeHeader Then
        Dim iCol As Long
        For iCol = 1 To rs.Fields.Count
            r.Cells(1, iCol).Value = rs.Fields(iCol - 1).Name
        Next
        Set r = r.Offset(1)
    End If
    r.CopyFromRecordset rs
    rs.Close
End Sub

Private Sub NonSelectQuery(ByVal queryName As String)
    CurrentDb.QueryDefs(queryName).Execute dbFailOnError
End Sub

Private Function SelectSqls(sqls As Collection) As Variant
    Dim sql As String
    sql = sqls(1)
    Dim i As Long
    For i = 2 To sqls.Count
        sql = sql & " UNION ALL " & sqls(i)
    Next
    With CurrentDb.OpenRecordset(sql)
        .MoveLast
        Dim rc As Long
        rc = .RecordCount
        .MoveFirst
        SelectSqls = .GetRows(rc)
        .Close
    End With
End Function

Private Function SelectSql(ByVal sql As String) As Variant
    With CurrentDb.OpenRecordset(SQL)
        .MoveLast
        Dim rc As Long
        rc = .RecordCount
        .MoveFirst
        SelectSql = .GetRows(rc)
        .Close
    End With
End Function

Private Sub NonSelectSqls(sqls As Collection)
    DBEngine.BeginTrans
    Dim sql As Variant
    For Each sql In sqls
        CurrentDb.Execute sql, dbFailOnError
    Next
    DBEngine.CommitTrans
End Sub

Private Sub NonSelectSql(ByVal sql As String)
    CurrentDb.Execute sql, dbFailOnError
End Sub

'Usage: UpdateColumn "Table1", "Column1 = " & Quote("foo") & " And Column2 = " & Quote("bar"), "Column3", Quote("baz")
Private Sub UpdateColumn(ByVal table As String, ByVal key As String, ByVal columnToUpdate As String, ByVal valueToSet As String)
    With CurrentDb.OpenRecordset(table)
        .FindFirst key
        .Edit
        .Fields(columnToUpdate).value = valueToSet
        .Update
    End With
End Sub

private Function HasRecord(ByVal table As String, ByVal where As String) As Boolean
    With CurrentDb.OpenRecordset("SELECT COUNT(*) FROM " & table & " WHERE " & where)
        Dim count As Variant
        count = .GetRows 'Seemingly redundant but necessary because .GetRows(0, 0) makes an error
        HasRecord = (0 < count(0, 0))
        .Close
    End With
End Function

Private Function InsertOrUpdate(ByVal table As String, ByVal pkName As String, ByVal pkValue As String, ByVal csv As String) As Boolean
    With CurrentDb.OpenRecordset("SELECT COUNT(*) FROM " & table & " WHERE " & pkName & " = " & pkValue)
        Dim count As Variant
        count = .GetRows 'Seemingly redundant but necessary because .GetRows(0, 0) makes an error
        If 0 < count(0, 0) Then CurrentDb.Execute "DELETE FROM " & table & " WHERE " & pkName & " = " & pkValue, dbFailOnError
        .Close
    End With
    CurrentDb.Execute "INSERT INTO " & table & " VALUES(" & csv & ")"
End Function

Private Sub ImportObject(ByVal type1 As AcObjectType, ByVal name As String)
    Const IMPORT_FROM As String = "release.mdb"
    On Error Resume Next
    DoCmd.DeleteObject type1, name
    On Error GoTo 0
    DoCmd.TransferDatabase acImport, "Microsoft Access", CurrentProject.Path & "\" & IMPORT_FROM, type1, name, name
End Sub

Private Sub ExportObject(ByVal type1 As AcObjectType, ByVal name As String)
    Const EXPORT_TO As String = "production.mdb"
    DoCmd.TransferDatabase acExport, "Microsoft Access", CurrentProject.Path & "\" & EXPORT_TO, type1, name, name
End Sub

Private Sub ReleaseByImporting()
    ImportObject acForm, "form1"
    ImportObject acModule, "module1"
    ImportObject acQuery, "query1"
    ImportObject acTable, "table1"
End Sub

Private Sub ReleaseByExporting()
    Dim ao As AccessObject
    For Each ao In CurrentProject.AllForms
        DoCmd.Close acForm, ao.name, acSaveNo
    Next
    ExportObject acForm, "form1"
    ExportObject acModule, "module1"
    ExportObject acQuery, "query1"
    ExportObject acTable, "table1"
End Sub

Private Sub EnableCloseMinMaxButtons()
    Dim ao As AccessObject
    Dim v
    For Each ao In CurrentProject.AllForms
        DoCmd.Close acForm, ao.Name, acSaveNo
        DoCmd.OpenForm ao.Name, acDesign, windowmode:=acHidden
        With Forms(ao.Name)
            .CloseButton = True
            .MinMaxButtons = 3 'http://msdn.microsoft.com/en-us/library/office/ff845417.aspx
        End With
        DoCmd.Close acForm, ao.Name, acSaveYes
    Next
End Sub

Private Sub SetFormat()
    On Error Resume Next
    Const TWIPS As Long = 567
    Const LABEL_WIDTH As Long = 2 * TWIPS
    Const TEXTBOX_COMBOBOX_WIDTH As Long = 4.5 * TWIPS
    Const LABEL_TEXTBOX_COMBOBOX_HEIGHT As Long = 0.5 * TWIPS
    Dim ao As AccessObject
    For Each ao In CurrentProject.AllForms
        DoCmd.OpenForm ao.Name, acDesign
        Dim c As Control
        For Each c In Forms(ao.Name).Controls
            c.FontSize = 10
            c.FontName = "Meiryo UI"
            c.ForeColor = RGB(0, 0, 0)
            c.AllowAutoCorrect = False
            c.AutoExpand = False
            Select Case c.ControlType
                Case acLabel
                c.Width = LABEL_WIDTH
                c.Height = LABEL_TEXTBOX_COMBOBOX_HEIGHT
                c.InSelection = True
                DoCmd.RunCommand acCmdBringToFront
                c.InSelection = False
                Case acTextBox
                c.Width = TEXTBOX_COMBOBOX_WIDTH
                c.Height = LABEL_TEXTBOX_COMBOBOX_HEIGHT
                Case acComboBox
                c.Width = TEXTBOX_COMBOBOX_WIDTH
                c.Height = LABEL_TEXTBOX_COMBOBOX_HEIGHT
            End Select
        Next
    Next
    On Error GoTo 0
End Sub

Private Function TwipToCm(ByVal twip As Double) As Double
    Const CM_PER_TWIP As Double = 1 / 567
    TwipToCm = twip * CM_PER_TWIP
End Function

Private Function CmToTwip(ByVal cm As Double) As Double
    Const TWIP_PER_CM As Double = 567
    CmToTwip = cm * TWIP_PER_CM
End Function