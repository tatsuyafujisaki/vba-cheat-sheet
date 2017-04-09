Option Explicit

Private Sub PrintRecordset(rs As Recordset)
    Debug.Print GetFieldNames(rs)
    Debug.Print GetRows(rs)
End Sub

Private Function GetRows(rs As DAO.Recordset) As String
    GetRows = ""
    With rs
        Dim table As Variant
        table = .GetRows(GetRecordCount(rs))
        Dim iRow As Long
        For iRow = 0 To UBound(table, 2)
            GetRows = GetRows & table(0, iRow)
            Dim iCol As Long
            For iCol = 1 To UBound(table)
                GetRows = GetRows & ", " & table(iCol, iRow)
            Next
            GetRows = GetRows & vbCrLf
        Next
    End With
End Function

Private Function GetFieldNames(rs As DAO.Recordset) As String
    With rs
        GetFieldNames = .Fields(0).name
        Dim iCol As Long
        For iCol = 1 To .Fields.Count - 1
            GetFieldNames = GetFieldNames & ", " & .Fields(iCol).name
        Next
    End With
End Function

Private Function GetRecordCount(rs As DAO.Recordset) As Long
    With rs
        .MoveLast
        GetRecordCount = .RecordCount
        .MoveFirst
    End With
End Function

Private Sub PrintVBComponents()
    Dim vbcs As Object
    Set vbcs = CreateObject("System.Collections.ArrayList")
    Dim vbc As Variant
    For Each vbc In VBE.VBProjects(1).VBComponents 'NB: VBE.ActiveProject is not always available
       vbcs.Add vbc.name 'vbc is VBComponent here
    Next
    vbcs.Sort
    For Each vbc In vbcs
        Debug.Print vbc 'vbc is String here
    Next
End Sub