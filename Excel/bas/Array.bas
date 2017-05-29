Option Explicit

Private Function GetArray(ByVal xs As Variant) As Variant
    GetArray = IIf(IsArray(xs), xs, Array(xs))
End Function

Private Function GetUpdateStatement(ByVal table As String, ByVal columns As Variant, ByVal values As Variant, ByVal where As String) As String
    GetUpdateStatement = "UPDATE " & table & " SET " & Join(GetMixedArray(columns, values, "="), ",") & " WHERE " & where
End Function

Private Function GetConcatenated(ByVal xs As Variant) As String
    GetConcatenated = vbNullString
    Dim x As Variant
    For Each x In xs
        GetConcatenated = GetConcatenated & x & vbLf
    Next
    GetConcatenated = Left$(GetConcatenated, Len(GetConcatenated) - 1)
End Function

Private Function GetMixedArray(ByVal xs1 As Variant, ByVal xs2 As Variant, ByVal delimiter As String) As Variant
    ReDim xs(LBound(xs1) To UBound(xs1))
    Dim i As Long
    For i = LBound(xs1) To UBound(xs1)
        xs(i) = xs1(i) & delimiter & xs2(i)
    Next
    GetMixedArray = xs
End Function

Private Function InArray(ByVal xs As Variant, ByVal findMe As String) As Boolean
    InArray = False
    Dim x As Variant
    For Each x In xs
        If StrComp(x, findMe, vbTextCompare) = 0 Then
            InArray = True
            Exit For
        End If
    Next
End Function

Private Function GetDimensionCount(ByVal xs As Variant) As Long
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

Private Function ArrayToTable(ByVal xs As Variant) As Variant
    Dim bounds As Variant
    bounds = Array(LBound(xs), UBound(xs))

    ReDim table(bounds(0) To bounds(0), bounds(0) To bounds(1))
    Dim columnIndex As Long
    For columnIndex = bounds(0) To bounds(1)
        table(bounds(0), columnIndex) = xs(columnIndex)
    Next
    ArrayToTable = table
    Erase xs
End Function

Private Function GetMinLong(ByVal xs As Variant, ByVal columnIndex As Long) As Long
    GetMinLong = 2147483647 'Highest possible long
    Dim rowIndex As Long
    For rowIndex = LBound(xs) To UBound(xs)
        If IsNumeric(xs(rowIndex, columnIndex)) And (xs(rowIndex, columnIndex) < GetMinLong) Then GetMinLong = xs(rowIndex, columnIndex)
    Next
End Function

Private Function GetMaxLong(ByVal xs As Variant, ByVal columnIndex As Long) As Long
    GetMaxLong = -2147483648# 'Lowest possible long
    Dim rowIndex As Long
    For rowIndex = LBound(xs) To UBound(xs)
        If IsNumeric(xs(rowIndex, columnIndex)) And (GetMaxLong < xs(rowIndex, columnIndex)) Then GetMaxLong = xs(rowIndex, columnIndex)
    Next
End Function

Private Function GetMinMaxLongs(ByVal xs As Variant, ByVal columnIndex As Long) As Variant
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

Private Function GetMinDate(ByVal xs As Variant, ByVal columnIndex As Long) As Date
    GetMinDate = DateSerial(9999, 12, 31)
    Dim rowIndex As Long
    For rowIndex = LBound(xs) To UBound(xs)
        If IsDate(xs(rowIndex, columnIndex)) And (xs(rowIndex, columnIndex) < GetMinDate) Then GetMinDate = xs(rowIndex, columnIndex)
    Next
End Function

Private Function GetMaxDate(ByVal xs As Variant, ByVal columnIndex As Long) As Date
    GetMaxDate = DateSerial(100, 1, 1)
    Dim rowIndex As Long
    For rowIndex = LBound(xs) To UBound(xs)
        If IsDate(xs(rowIndex, columnIndex)) And (GetMaxDate < xs(rowIndex, columnIndex)) Then GetMaxDate = xs(rowIndex, columnIndex)
    Next
End Function

Private Function GetMinMaxDates(ByVal xs As Variant, ByVal columnIndex As Long) As Variant
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
End Function
