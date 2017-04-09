Option Explicit

Private Function GetArray(xs)
    GetArray = IIf(IsArray(xs), xs, Array(xs))
End Function

Private Function GetUpdateStatement(ByVal table As String, columns, values, ByVal where As String) As String
    GetUpdateStatement = "UPDATE " & table & " SET " & Join(GetMixedArray(columns, values, "="), ",") & " WHERE " & where
End Function

Private Function GetConcatenated(xs) As String
    GetConcatenated = ""
    Dim e
    For Each e In xs
        GetConcatenated = GetConcatenated & e & vbLf
    Next
    GetConcatenated = Left(GetConcatenated, Len(GetConcatenated) - 1)
End Function

Private Function GetMixedArray(xs1, xs2, ByVal delimiter As String)
    ReDim xs(LBound(xs1) To UBound(xs1))
    Dim i As Long
    For i = LBound(xs1) To UBound(xs1)
        xs(i) = xs1(i) & delimiter & xs2(i)
    Next
    GetMixedArray = xs
End Function

Private Function InArray(xs, ByVal findMe As String) As Boolean
    InArray = False
    Dim element
    For Each element In xs
        If StrComp(element, findMe, vbTextCompare) = 0 Then
            InArray = True
            Exit For
        End If
    Next
End Function

Private Function GetDimensionCount(xs) As Long
    Dim i As Long
    i = 1
    On Error Resume Next
    Do
        Dim devnull As Long
        devnull = UBound(xs, i)
        i = i + 1
    Loop While Err.Number = 0
    On Error GoTo 0
    GetDimensionCount = i - 2
End Function

Private Function ArrayToTable(xs)
    Dim bounds
    bounds = Array(LBound(xs), UBound(xs))

    ReDim table(bounds(0) To bounds(0), bounds(0) To bounds(1))
    Dim columnIndex As Long
    For columnIndex = bounds(0) To bounds(1)
        table(bounds(0), columnIndex) = xs(columnIndex)
    Next
    ArrayToTable = table
    Erase xs
End Function

Private Function GetMinLong(xs, ByVal columnIndex As Long) As Long
    GetMinLong = 2147483647 'Highest possible long
    Dim rowIndex As Long
    For rowIndex = LBound(xs) To UBound(xs)
        If IsNumeric(xs(rowIndex, columnIndex)) And (xs(rowIndex, columnIndex) < GetMinLong) Then GetMinLong = xs(rowIndex, columnIndex)
    Next
End Function

Private Function GetMaxLong(xs, ByVal columnIndex As Long) As Long
    GetMaxLong = -2147483648# 'Lowest possible long
    Dim rowIndex As Long
    For rowIndex = LBound(xs) To UBound(xs)
        If IsNumeric(xs(rowIndex, columnIndex)) And (GetMaxLong < xs(rowIndex, columnIndex)) Then GetMaxLong = xs(rowIndex, columnIndex)
    Next
End Function

Private Function GetMinMaxLongs(xs, ByVal columnIndex As Long)
    Dim min As Long
    min = 2147483647 'Highest possible long
    Dim max As Long
    max = -2147483648# 'Lowest possible long
    Dim rowIndex As Long
    For rowIndex = LBound(xs) To UBound(xs)
        If IsNumeric(xs(rowIndex, columnIndex)) Then
            If xs(rowIndex, columnIndex) < min Then
                min = xs(rowIndex, columnIndex)
            ElseIf max < xs(rowIndex, columnIndex) Then
                max = xs(rowIndex, columnIndex)
            End If
        End If
    Next
    GetMinMaxLongs = Array(min, max)
End Function

Private Function GetMinDate(xs, ByVal columnIndex As Long) As Date
    GetMinDate = DateSerial(9999, 12, 31)
    Dim rowIndex As Long
    For rowIndex = LBound(xs) To UBound(xs)
        If IsDate(xs(rowIndex, columnIndex)) And (xs(rowIndex, columnIndex) < GetMinDate) Then GetMinDate = xs(rowIndex, columnIndex)
    Next
End Function

Private Function GetMaxDate(xs, ByVal columnIndex As Long) As Date
    GetMaxDate = DateSerial(100, 1, 1)
    Dim rowIndex As Long
    For rowIndex = LBound(xs) To UBound(xs)
        If IsDate(xs(rowIndex, columnIndex)) And (GetMaxDate < xs(rowIndex, columnIndex)) Then GetMaxDate = xs(rowIndex, columnIndex)
    Next
End Function

Private Function GetMinMaxDates(xs, ByVal columnIndex As Long)
    Dim min As Date
    min = DateSerial(9999, 12, 31)

    Dim max As Date
    max = DateSerial(100, 1, 1)

    Dim rowIndex As Long
    For rowIndex = LBound(xs) To UBound(xs)
        If IsDate(xs(rowIndex, columnIndex)) Then
            If xs(rowIndex, columnIndex) < min Then
                min = xs(rowIndex, columnIndex)
            ElseIf max < xs(rowIndex, columnIndex) Then
                max = xs(rowIndex, columnIndex)
            End If
        End If
    Next
    GetMinMaxDates = Array(min, max)
End Function