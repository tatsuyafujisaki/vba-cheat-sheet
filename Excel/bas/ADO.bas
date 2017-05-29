Option Explicit

Private Function GetConnectionString(ByVal isProduction As Boolean) As String
    Const productionCs As String = "Driver=Adaptive Server Enterprise;Server=server1;Port=port1;Db=db1;Uid=uid1;Pwd=pwd1"
    Const developmentCs As String = "Driver=Adaptive Server Enterprise;Server=server1;Port=port1;Db=db1;Uid=uid1;Pwd=pwd1"
    GetConnectionString = IIf(isProduction, productionCs, developmentCs)
End Function

Private Function SelectSql(ByVal sql As String) As Variant
    Const connectionString As String = "Driver=SQL Server;Server=server1,port1;Database=database1;Uid=uid1;Pwd=pwd1"
    Dim cn As New ADODB.Connection 'Microsoft ActiveX Data Object x.x Library
    cn.Open connectionString
    With New ADODB.Recordset
        .Open sql, cn, adOpenStatic
        If .RecordCount = 0 Then 'IIf makes an error
            SelectSql = Null
        Else
            SelectSql = WorksheetFunction.Transpose(.GetRows)
        End If
        .Close
    End With
    cn.Close
End Function

Private Sub NonSelectSqls(ByVal sqls As Collection)
    Const connectionString As String = "Driver=SQL Server;Server=server1,port1;Database=database1;Uid=uid1;Pwd=pwd1"
    With New ADODB.Connection 'Microsoft ActiveX Data Object x.x Library
        .Open connectionString
        .BeginTrans
        Dim sql As Object
        For Each sql In sqls
            .Execute sql
        Next
        .CommitTrans
        .Close
    End With
End Sub

Private Sub NonSelectSql(ByVal sql As String)
    Const connectionString As String = "Driver=SQL Server;Server=server1,port1;Database=database1;Uid=uid1;Pwd=pwd1"
    With New ADODB.Connection 'Microsoft ActiveX Data Object x.x Library
        .Open connectionString
        .Execute sql
        .Close
    End With
End Sub
