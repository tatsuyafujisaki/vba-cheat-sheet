Option Explicit

Private Sub CompareTwoSheets()
    Dim sourceWs1 As Worksheet
    Set sourceWs1 = ThisWorkbook.Worksheets("Source1")

    Dim sourceWs2 As Worksheet
    Set sourceWs2 = ThisWorkbook.Worksheets("Source2")

    Dim destinationWs1 As Worksheet
    Set destinationWs1 = ThisWorkbook.Worksheets("Destination")

    Dim rowCount As Long
    rowCount = WorksheetFunction.min(sourceWs1.UsedRange.rows.Count, sourceWs2.UsedRange.rows.Count)

    Dim columnCount As Long
    columnCount = WorksheetFunction.min(sourceWs1.UsedRange.columns.Count, sourceWs2.UsedRange.columns.Count)

    Dim sourceRange1 As Range
    Set sourceRange1 = sourceWs1.Cells(1, 1).Resize(rowCount, columnCount)

    Dim sourceRange2 As Range
    Set sourceRange2 = sourceWs2.Cells(1, 1).Resize(rowCount, columnCount)

    Dim destinationRange As Range
    Set destinationRange = destinationWs1.Cells(1, 1).Resize(rowCount, columnCount)

    destinationWs1.Cells.Clear

    Dim i As Long
    For i = 1 To columnCount Step 3
        destinationRange.columns(i).Value = sourceRange1.columns(i).Value
        destinationRange.columns(i + 1).Value = sourceRange2.columns(i).Value
        destinationRange.columns(i + 2).Value = "=RC[-2] = RC[-1]"
        HighlightTrueFalse destinationRange.columns(i + 2)
    Next
End Sub

Private Sub HighlightNamedRange()
  Dim n As Name
  For Each n In ThisWorkbook.Names
    Dim v
    v = Split(n.Value, "!")
    ThisWorkbook.Sheets(Replace(v(0), "=", "")).Range(v(1)).Interior.ColorIndex = 6 'Yellow
  Next
End Sub

Private Sub HighlightTrueFalse(r As Range)
    With r.FormatConditions
        .Delete
        With .Add(xlCellValue, xlEqual, "TRUE")
            .Font.Bold = True
            .Font.ThemeColor = xlThemeColorDark1
            .Interior.ThemeColor = xlThemeColorAccent1
        End With
        With .Add(xlCellValue, xlEqual, "FALSE")
            .Font.Bold = True
            .Font.ThemeColor = xlThemeColorDark1
            .Interior.ThemeColor = xlThemeColorAccent2
        End With
    End With
End Sub

Private Sub HighlightDuplicates(ws As Worksheet)
    Const columnIndex As Long = 1
    Dim col As Range
    Set col = Intersect(ws.UsedRange, ws.columns(columnIndex).EntireColumn)
    ws.UsedRange.Interior.ColorIndex = xlColorIndexNone
    Dim singleCell As Range
    For Each singleCell In col
        If 1 < WorksheetFunction.CountIf(col, singleCell.Value) Then singleCell.Interior.ColorIndex = 6 'Yellow
    Next
End Sub

Private Sub HighlightMatched(ws As Worksheet)
    Const sourceColumn1Index As Long = 1
    Const sourceColumn2Index As Long = 2
    ws.UsedRange.Interior.ColorIndex = xlColorIndexNone

    Dim col1 As Range
    Set col1 = GetColumn(ws, sourceColumn1Index)

    Dim col2 As Range
    Set col2 = GetColumn(ws, sourceColumn2Index)

    ToNumeric Union(col1, col2)
    HighlightMatchedCallback col1, col2.Value
    HighlightMatchedCallback col2, col1.Value
End Sub

Private Sub HighlightMatchedCallback(col1 As Range, col2)
    Dim singleCell As Range
    For Each singleCell In col1
        If Not IsEmpty(singleCell.Value) Then
            On Error Resume Next
            WorksheetFunction.VLookup singleCell.Value, col2, 1, False
            If Err.Number = 0 Then singleCell.Interior.ColorIndex = 6 'Yellow
            On Error GoTo 0
        End If
    Next
End Sub

Private Sub FindMatched(ws As Worksheet)
    Const sourceColumn1Index As Long = 1
    Const sourceColumn2Index As Long = 2
    Const destinationColumnIndex As Long = 3

    Dim col1 As Range
    Set col1 = GetColumn(ws, sourceColumn1Index)

    Dim col2 As Range
    Set col2 = GetColumn(ws, sourceColumn2Index)

    ToNumeric Union(col1, col2)

    Dim matched As New Dictionary
    matched.CompareMode = TextCompare

    FindMatchedCallback matched, col1, col2.Value
    FindMatchedCallback matched, col2, col1.Value
    PasteDictionary matched, ws.columns(destinationColumnIndex)
End Sub

Private Sub FindMatchedCallback(d As Dictionary, col1 As Range, col2)
    Dim singleCell As Range
    For Each singleCell In col1
        If Not IsEmpty(singleCell.Value) Then
            On Error Resume Next
            WorksheetFunction.VLookup singleCell.Value, col2, 1, False
            If Err.Number = 0 And Not d.Exists(singleCell.Value) Then d.Add singleCell.Value, singleCell.Value
            On Error GoTo 0
        End If
    Next
End Sub

Private Sub PasteDictionary(d As Dictionary, r As Range)
    ReDim table(d.Count, 0)
    Dim i As Long
    For i = 0 To d.Count - 1
        table(i, 0) = d.Items(i)
    Next
    PasteTable r, table
    Erase table
End Sub

Private Sub ToNumeric(r As Range)
    Dim c As Range
    For Each c In r
        If c.Value <> "" And IsNumeric(c.Value) Then c.Value = CDbl(c.Value)
    Next
End Sub

Private Function GetColumn(ws As Worksheet, ByVal colIndex As Long) As Range
    Set GetColumn = ws.Range(Cells(IIf(ws.Cells(1, colIndex).Value <> "", 1, ws.Cells(1, colIndex).End(xlDown).row), colIndex), Cells(ws.Cells(ws.rows.Count, colIndex).End(xlUp).row, colIndex))
End Function

Private Sub PasteTable(r As Range, table)
      r.Resize(UBound(table) - LBound(table) + 1, UBound(table, 2) - LBound(table, 2) + 1) = table
End Sub