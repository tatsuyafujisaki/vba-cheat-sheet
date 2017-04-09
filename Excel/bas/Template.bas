Option Explicit

Private Sub SetFont()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        With ws.Cells.Font
            .Name = "Meiryo UI"
            .Size = 10
        End With
    Next
End Sub

Private Sub SetCursorToA1()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        With ws
            .Activate
            .Cells(1, 1).Select
        End With
    Next
    ThisWorkbook.Worksheets(1).Activate
End Sub

Private Function AreDates(xs) As Boolean
    AreDates = True
    Dim x
    For Each x In xs
        If (x <> "") And Not IsDate(x) Then
            AreDates = False
            Exit For
        End If
    Next
End Function

Private Function AreNumeric(xs) As Boolean
    AreNumeric = True
    Dim x
    For Each x In xs
        If (x <> "") And Not IsNumeric(x) Then
            AreNumeric = False
            Exit For
        End If
    Next
End Function