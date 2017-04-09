Option Explicit

Private Sub LocalizeLinkTable(ByVal remoteTable As String, ByVal localTable As String)
    Const CONNECT As String = "ODBC;DRIVER=SQL Server;SERVER=server1,port1;DATABASE=database1;UID=uid1;PWD=pwd1"
    On Error Resume Next
    DoCmd.DeleteObject acTable, localTable
    On Error GoTo 0
    DoCmd.TransferDatabase acImport, "ODBC Database", CONNECT, acTable, remoteTable, localTable
End Sub

Private Sub PrintConnectionStrings()
    Dim td As DAO.TableDef
    For Each td In CurrentDb.TableDefs
        If Len(td.Connect) Then Debug.Print td.name & vbCrLf & td.Connect & vbCrLf
    Next
End Sub

Private Sub SetPassThroughQueryConnectionString()
    Const PRD As String = "ODBC;DRIVER=SQL Server;SERVER=server1,port1;DATABASE=database1;UID=uid1;PWD=pwd1"
    Const DEV As String = "ODBC;DRIVER=SQL Server;SERVER=server1,port1;DATABASE=database1;UID=uid1;PWD=pwd1"
    Dim cs As String
    cs = IIf(IsProduction, PRD, DEV)
    Dim qd As DAO.QueryDef
    For Each qd In CurrentDb.QueryDefs
        If qd.Type = dbQSQLPassThrough Then qd.Connect = cs
    Next
End Sub

Private Sub UpdateLocalLinkTables()
    Dim findMeList As Variant
    findMeList = Array("Accdb1.accdb", "Mdb1.mdb")
    Dim td As DAO.TableDef
    For Each td In CurrentDb.TableDefs
        Dim path As String
        path = GetDbPath(td.CONNECT)
        If path <> "" Then
            Dim findMe As Variant
            For Each findMe In findMeList
                If Right(path, Len(findMe)) = findMe Then
                    td.CONNECT = ";Database=" & CurrentProject.path & "\Server\" & findMe
                    td.RefreshLink
                    Exit For
                End If
            Next
        End If
    Next
End Sub

Private Function GetDbPath(ByVal connectionString As String) As String
    With New RegExp 'Microsoft VBScript Regular Expressions x.x
        .Pattern = "DATABASE=(.+\.(accdb|mdb))"
        Dim mc As MatchCollection
        Set mc = .Execute(connectionString)
        If mc.Count Then 'NB: IIf makes an error
            GetDbPath = mc(0).SubMatches(0)
        Else
            GetDbPath = ""
        End If
    End With
End Function