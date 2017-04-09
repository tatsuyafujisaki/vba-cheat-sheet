Option Explicit

Private Sub CreatePivotTableFromRange(src As Range, singleCell As Range, headers, cols, data)
    Const DATA_ITEM As Long = 0
    Const DATA_FUNC As Long = 1
    With ThisWorkbook.PivotCaches.Create(xlDatabase, src).CreatePivotTable(singleCell)
        .TableStyle2 = "PivotStyleLight8"
        .CompactLayoutRowHeader = rows(0)
        .CompactLayoutColumnHeader = cols(0)
        Dim item
        For Each item In rows
            .PivotFields(item).Orientation = xlRowField
        Next
        For Each item In cols
            .PivotFields(item).Orientation = xlColumnField
        Next
        With .PivotFields(data(DATA_ITEM))
            .Orientation = xlDataField
            .Function = data(DATA_FUNC)
        End With
        On Error Resume Next
        Dim pf As PivotField
        For Each pf In .PivotFields
            pf.PivotItems("(Blank)").Visible = False
        Next
        For Each pf In .PivotFields
            With pf
                .PivotItems("その他").Position = .PivotItems.Count
            End With
        Next
        On Error GoTo 0
        '.DataPivotField.PivotItems(1).Caption = "your custom data header"
    End With
End Sub

Private Sub CreatePivotTableFromFile(ByVal dir As String, ByVal file As String, singleCell As Range, headers, rows, cols, data)
    Const HEADER_ROW As Long = 0
    Const HEADER_COL As Long = 1
    Const HEADER_DATA As Long = 2
    Const DATA_ITEM As Long = 0
    Const DATA_FUNC As Long = 1
    With ThisWorkbook.PivotCaches.Create(xlExternal)
        .Connection = "ODBC;Driver={Microsoft Text Driver (*.txt; *.csv)};DBQ=" & dir
        .CommandText = "SELECT * FROM " & file
        With .CreatePivotTable(singleCell)
            .TableStyle2 = "PivotStyleLight8"
            .CompactLayoutRowHeader = headers(HEADER_ROW)
            .CompactLayoutColumnHeader = headers(HEADER_COL)
            Dim item
            For Each item In rows
                .PivotFields(item).Orientation = xlRowField
            Next
            For Each item In cols
                .PivotFields(item).Orientation = xlColumnField
            Next
            With .PivotFields(data(DATA_ITEM))
                .Orientation = xlDataField
                .Function = data(DATA_FUNC)
            End With
            On Error Resume Next
            Dim pf As PivotField
            For Each pf In .PivotFields
                pf.PivotItems("(Blank)").Visible = False
            Next
            For Each pf In .PivotFields
                With pf
                    .PivotItems("その他").Position = .PivotItems.Count
                End With
            Next
            On Error GoTo 0
            .DataPivotField.PivotItems(1).Caption = headers(HEADER_DATA)
        End With
    End With
End Sub