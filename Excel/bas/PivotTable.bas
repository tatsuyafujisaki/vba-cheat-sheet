Option Explicit

Private Sub CreatePivotTableFromRange(ByVal src As Range, ByVal singleCell As Range, ByVal columns As Variant, ByVal data As Variant)
    Const DATA_ITEM As Long = 0
    Const DATA_FUNC As Long = 1
    With ThisWorkbook.PivotCaches.Create(xlDatabase, src).CreatePivotTable(singleCell)
        .TableStyle2 = "PivotStyleLight8"
        .CompactLayoutRowHeader = rows(0)
        .CompactLayoutColumnHeader = columns(0)
        Dim item As Variant
        For Each item In rows
            .PivotFields(item).Orientation = xlRowField
        Next
        For Each item In columns
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

Private Sub CreatePivotTableFromFile(ByVal dir As String, ByVal file As String, ByVal singleCell As Range, ByVal headers As Variant, ByVal rows As Variant, ByVal columns As Variant, ByVal data As Variant)
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
            Dim item As Variant
            For Each item In rows
                .PivotFields(item).Orientation = xlRowField
            Next
            For Each item In columns
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
