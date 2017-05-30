Option Explicit

Private Function TableToCSV(table) As String
    TableToCSV = ""
    Dim bounds
    bounds = Array(LBound(table, 2), UBound(table, 2))
    Dim rowIndex As Long
    For rowIndex = LBound(table) To UBound(table)
        TableToCSV = TableToCSV & vbCrLf & table(rowIndex, bounds(0))
        Dim columnIndex As Long
        For columnIndex = bounds(0) + 1 To bounds(1)
            TableToCSV = TableToCSV & Chr(44) & table(rowIndex, columnIndex)
        Next
    Next
    TableToCSV = Mid(TableToCSV, 3)
End Function

Private Function CsvToInserts(ByVal csvPath As String) As Collection
    Dim sqls As New Collection
    Dim csv
    csv = ReadCSV(csvPath)
    With New FileSystemObject
        Dim tableName As String
        tableName = .GetBaseName(csvPath)
    End With
    sqls.Add "DELETE FROM " & tableName
    Dim nCols As Long
    nCols = UBound(csv, 2)
    Dim rowIndex As Long
    For rowIndex = 1 To UBound(csv)
        Dim columnIndex As Long
        Dim values As String
        values = ""
        For columnIndex = 1 To nCols
            values = values & "," & Quote(csv(rowIndex, columnIndex))
        Next
        sqls.Add "INSERT INTO " & tableName & " VALUES(" & Mid(values, 2) & ")"
    Next
    Set CsvToInserts = sqls
End Function

Private Function ExcelToCsv(ByVal excelPath As String, Optional ByVal iSheet As Long = 1) As String
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
    End With
    With New FileSystemObject
        Dim csvPath As String
        csvPath = .BuildPath(.GetParentFolderName(excelPath), .GetBaseName(excelPath)) & ".csv"
    End With
    With Workbooks.Open(excelPath, ReadOnly:=True)
        .Sheets(iSheet).Activate
        .SaveAs csvPath, xlCSV
        .Close
    End With
    With Application
        .DisplayAlerts = True
        .ScreenUpdating = True
    End With
    ExcelToCsv = csvPath
End Function

Private Sub SheetToCSV(ByVal path As String)
    Application.DisplayAlerts = False
    Me.SaveAs path, xlCSV
    Application.DisplayAlerts = True
End Sub

Private Sub CsvToSheet(ByVal path As String)
    With Me.QueryTables.Add("TEXT;" & path, Me.Cells(1, 1))
        .TextFileCommaDelimiter = True
        .Refresh
        .Delete
    End With
End Sub

Private Function ReadText(ByVal path As String) As String
    With New FileSystemObject
        With .GetFile(path).OpenAsTextStream
            ReadText = GetTailingNewLinesRemoved(GetUnifiedNewLines(.ReadAll))
            .Close
        End With
    End With
End Function

Private Function ReadCSV(ByVal path As String)
    Dim rows
    rows = Split(ReadText(path), vbLf)

    Dim nRows As Long
    nRows = UBound(rows)

    Dim row
    row = Split(rows(0), ",")

    Dim nCols As Long
    nCols = UBound(row)

    ReDim table(nRows, nCols)
    Dim rowIndex As Long
    For rowIndex = 0 To nRows
        row = Split(rows(rowIndex), ",")
        Dim columnIndex As Long
        For columnIndex = 0 To nCols
            table(rowIndex, columnIndex) = row(columnIndex)
        Next
    Next
    ReadCSV = table
End Function

Private Function ReadCsvBySheet(ByVal path As String)
    Application.ScreenUpdating = False
    With Workbooks.Open(path, ReadOnly:=True)
        ReadCSV = .Sheets(1).UsedRange
        .Close
    End With
    Application.ScreenUpdating = True
End Function

Private Function ReadCsvByTextDriver(ByVal dir As String, ByVal file As String, Optional ByVal hasHeader As Boolean = False)
    Const CONNECTION_STRING As String = "Driver={Microsoft Text Driver (*.txt; *.csv)};DBQ="

    Dim sf As New SchemaFactory
    sf.Init dir, file, hasHeader

    Dim cn As New ADODB.Connection        'Microsoft ActiveX Data Object x.x Library
    cn.Open CONNECTION_STRING & dir
    With New ADODB.Recordset
        .Open "SELECT * FROM " & file, cn
        ReadCsvByTextDriver = TransposeTable(.GetRows)
        .Close
    End With
    cn.Close
End Function

'This is better than the other for two reasons
'-> The other needs Microsoft Scripting Runtime in References
'-> The other creates an object (however the creation time is negligible)
Private Sub WriteText(ByVal path As String, ByVal s As String)
    Dim fn As Integer
    fn = FreeFile

    Open path For Output As #fn
    Print #fn, s
    Close fn
End Sub

Private Sub WriteText(ByVal path As String, ByVal s As String)
    With New FileSystemObject
        'Set TristateTrue as 4th parameter for Unicode. Default is TristateFalse (Ascii)
        With .OpenTextFile(path, ForWriting, True)
            .Write s
            .Close
        End With
    End With
End Sub
