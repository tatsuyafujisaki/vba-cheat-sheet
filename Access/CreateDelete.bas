Option Explicit

Private Sub CallCreateTable()
    Dim fields As New Dictionary
    fields.Add "Field1", dbText
    fields.Add "Field2", dbText
    CreateTable "Table1", "Field1", fields
End Sub

Private Sub CreateTable(ByVal table As String, ByVal pk As String, fields As Dictionary)
    If TableExists(table) Then CurrentDb.TableDefs.Delete table
    Dim td As DAO.TableDef
    Set td = CurrentDb.CreateTableDef(table)
    Dim field As Variant
    For Each field In fields.Keys
        td.fields.Append td.CreateField(field, fields(field))
    Next
    Dim index As DAO.index
    Set index = td.CreateIndex("PrimaryKey")
    index.fields.Append index.CreateField(pk)
    index.Primary = True
    td.Indexes.Append index
    CurrentDb.TableDefs.Append td
End Sub

Private Sub CallAddRecords()
    ReDim table(3, 1) As String
    table(0, 0) = "Data at Row1 Column1"
    table(0, 1) = "Data at Row2 Column2"
    table(1, 0) = "Data at Row1 Column1"
    table(1, 1) = "Data at Row2 Column2"
    AddRecords "Table1", table
End Sub

Private Sub AddRecords(ByVal tableName As String, table As Variant)
    With CurrentDb.TableDefs(tableName).OpenRecordset
        Dim iRow As Long
        For iRow = 0 To UBound(table)
            .AddNew
            Dim iCol As Long
            For iCol = 0 To UBound(table, 2)
                .fields(iCol) = table(iRow, iCol)
            Next
            .Update
        Next
    End With
End Sub

Private Sub CreatePassThroughQuery(ByVal name As String, ByVal sql As String)
    On Error Resume Next
    CurrentDb.QueryDefs.Delete name
    On Error GoTo 0
    With CurrentDb.CreateQueryDef(name)
        .CONNECT = "ODBC;Driver=SQL Server;Server=server1,port1;Database=database1;Uid=uid1;Pwd=pwd1"
        .SQL = sql
        .Close
    End With
End Sub

Private Sub UnlinkTables()
    Dim td As DAO.TableDef
    For Each td In CurrentDb.TableDefs
        If td.connect <> "" Then CurrentDb.TableDefs.Delete td.name
    Next
End Sub

Private Sub LinkTable(ByVal connect As String, ByVal remoteName As String, ByVal localName As String)
    'connect = "ODBC;Driver=SQL Server;Server=server1;Database=database1;Uid=uid1;Pwd=pwd1"
    'connect = ";DATABASE=path/to/db.accde"
    On Error Resume Next
    CurrentDb.TableDefs.Delete localName
    On Error GoTo 0
    CurrentDb.TableDefs.Append CurrentDb.CreateTableDef(localName, dbAttachSavePWD, remoteName, connect)
End Sub

Private Sub CreateLocalQueries()
  Dim file As Variant
  For Each file In GetFiles(New Collection, CurrentProject.path & "\release\queries", Array("sql"))
    Dim name As String
    name = GetBaseName(file)
    On Error Resume Next
    CurrentDb.QueryDefs.Delete name
    On Error GoTo 0
    CurrentDb.CreateQueryDef name, ReadText(file)
  Next
End Sub

Private Sub DeleteObjectOfAllKinds(ByVal name As String)
    On Error Resume Next
    Dim AcObjectType As Long
    For AcObjectType = -1 To 12
        DoCmd.DeleteObject AcObjectType, name
    Next
    On Error GoTo 0
End Sub

Private Sub UnlinkTables()
    Dim findMeList As Variant
    findMeList = Array("Accdb1.accdb", "Mdb1.mdb")
    Dim td As DAO.TableDef
    For Each td In CurrentDb.TableDefs
        Dim path As String
        path = GetDbPath(td.connect)
        If path <> "" Then
            Dim findMe As Variant
            For Each findMe In findMeList
                If Right(path, Len(findMe)) = findMe Then
                    CurrentDb.TableDefs.Delete td.name
                    Exit For
                End If
            Next
        End If
    Next
End Sub

Private Sub RedirectLinkTables()
    Dim findMeList As Variant
    findMeList = Array("Accdb1.accdb", "Mdb1.mdb")
    Dim td As DAO.TableDef
    For Each td In CurrentDb.TableDefs
        Dim path As String
        path = GetDbPath(td.connect)
        If path <> "" Then
            Dim findMe As Variant
            For Each findMe In findMeList
                If Right(path, Len(findMe)) = findMe Then
                    td.connect = ";Database=" & CurrentProject.path & findMe
                    td.RefreshLink
                    Exit For
                End If
            Next
        End If
    Next
End Sub

Private Sub CompleteQuery(ByVal queryName As String, params As Dictionary)
    With New FileSystemObject
        With .GetFile(.BuildPath(CurrentProject.path, queryName & ".sql")).OpenAsTextStream
            Dim sql As String
            sql = .ReadAll
            .Close
        End With
    End With
    Dim k As Variant
    For Each k In params.Keys
        sql = Replace(sql, k, params(k))
    Next
    CurrentDb.QueryDefs(queryName).sql = sql
End Sub