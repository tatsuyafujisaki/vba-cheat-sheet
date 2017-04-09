Option Explicit

Private Function GetLastRowIndex(r As Range) As Long
    GetLastRowIndex = r.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
End Function

Private Function GetLastColumnIndex(r As Range) As Long
    GetLastColumnIndex = r.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
End Function

Private Function Move1(r As Range, ByVal down As Long, ByVal right As Long) As Range
    Set Move1 = r.offset(down, right)
End Function

'Top: Shrink if positive
'Left: Shrink if positive
'Bottom: Expand if positive
'Right: Expand if positive
Private Function Stretch(r As Range, ByVal top As Long, ByVal Left As Long, ByVal bottom As Long, ByVal right As Long) As Range
    Set Stretch = r.offset(top, Left).Resize(r.Rows.Count + bottom - top, r.Columns.Count + right - Left)
End Function

Private Function ExcludeHeader(r As Range) As Range
    Set ExcludeHeader = r.offset(1).Resize(r.Rows.Count - 1)
End Function

Private Function GetTable(r As Range)
    Dim r2 As Range
    Set r2 = r.CurrentRegion
    Dim offset As Long
    offset = r.Row - r2.Row
    GetTable = r2.offset(offset).Resize(r2.Rows.Count - offset).Value
End Function

Private Sub PasteTable(r As Range, table)
    r.Resize(UBound(table) - LBound(table) + 1, UBound(table, 2) - LBound(table, 2) + 1) = table
End Sub

Private Function GetDimensionCount(xs) As Long
    Dim i As Long
    i = 1
    On Error Resume Next
    Do
        Dim ignored As Long
        ignored = UBound(xs, i)
        i = i + 1
    Loop While Err.Number = 0
    On Error GoTo 0
    GetDimensionCount = i - 2
End Function

Private Function GetSlicedTable(table, ParamArray columnIndexOrString())
    ReDim subTable(1 To UBound(table), 1 To UBound(columnIndexOrString) + 1)
    Dim rowIndex As Long
    For rowIndex = 1 To UBound(subTable)
        Dim columnIndex As Long
        For columnIndex = 1 To UBound(subTable, 2)
            Dim v
            v = columnIndexOrString(columnIndex - 1)
            If VarType(v) = vbString Then 'IIf makes an error when v is column index beyond table
                subTable(rowIndex, columnIndex) = v
            Else
                subTable(rowIndex, columnIndex) = table(rowIndex, v)
            End If
        Next
    Next
    GetSlicedTable = subTable
End Function

Private Function MergeTables(table1, table2)
    Dim bounds1
    bounds1 = Array(LBound(table1), UBound(table1), LBound(table1, 2), UBound(table1, 2))

    Dim bounds2
    bounds2 = Array(LBound(table2), UBound(table2), LBound(table2, 2), UBound(table2, 2))

    Dim nRows1 As Long
    nRows1 = bounds1(1) - bounds1(0) + 1

    Dim nRows2 As Long
    nRows2 = bounds2(1) - bounds2(0) + 1

    Dim nCols1 As Long
    nCols1 = bounds1(3) - bounds1(2) + 1

    Dim nCols2 As Long
    nCols2 = bounds2(3) - bounds2(2) + 1

If nCols1 <> nCols2 Then Err.Raise 9 'Subscript out of range (https://support.microsoft.com/kb/146864)
    ReDim table(nRows1 + nRows2 - 1, nCols1 - 1)
    CopyTable table1, table, 0
    CopyTable table2, table, nRows1
    MergeTables = table
End Function

Private Sub CopyTable(table1, table2, ByVal rowIndex2 As Long)
    Dim bounds1: bounds1 = Array(LBound(table1), UBound(table1), LBound(table1, 2), UBound(table1, 2))
    Dim bounds2: bounds2 = Array(LBound(table2), UBound(table2), LBound(table2, 2), UBound(table2, 2))
    Dim nCols1 As Long: nCols1 = bounds1(3) - bounds1(2) + 1
    Dim nCols2 As Long: nCols2 = bounds2(3) - bounds2(2) + 1
If nCols1 <> nCols2 Then Err.Raise 9 'Subscript out of range (https://support.microsoft.com/kb/146864)
    Dim rowIndex1 As Long
    For rowIndex1 = bounds1(0) To bounds1(1)
        Dim columnIndex2 As Long
        columnIndex2 = 0

        Dim columnIndex1 As Long
        For columnIndex1 = bounds1(2) To bounds1(3)
            table2(rowIndex2, columnIndex2) = table1(rowIndex1, columnIndex1)
            columnIndex2 = columnIndex2 + 1
        Next
        rowIndex2 = rowIndex2 + 1
    Next
End Sub

'Use only when WorksheetFunction.Transpose makes Type Mismatch error
Private Function TransposeTable(table)
    Dim bounds
    bounds = Array(LBound(table, 2), UBound(table, 2), LBound(table), UBound(table))

    ReDim table2(bounds(0) To bounds(1), bounds(2) To bounds(3))
    Dim rowIndex As Long
    For rowIndex = bounds(0) To bounds(1)
        Dim columnIndex As Long
        For columnIndex = bounds(2) To bounds(3)
            table2(rowIndex, columnIndex) = table(columnIndex, rowIndex)
        Next
    Next
    TransposeTable = table2
End Function
