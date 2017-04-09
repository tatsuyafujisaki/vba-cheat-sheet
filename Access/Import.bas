Option Explicit

Private Sub ImportVBComponentsAndObjects(paths As Variant)
    Dim fso As New FileSystemObject
    Dim path As Variant
    For Each path In paths
        Select Case fso.GetExtensionName(path)
            Case "bas"
                ImportVBComponent path
            Case "cls"
                ImportVBComponent path
            Case "form"
                ImportObject acForm, path
            Case "report"
                ImportObject acReport, path
            Case Else
                Err.Raise 93 'Invalid pattern string (https://support.microsoft.com/kb/146864)
        End Select
    Next
End Sub

Private Sub ImportVBComponent(ByVal path As String)
    With New FileSystemObject
        Dim module As String
        module = .GetBaseName(path)
    End With
    With VBE.VBProjects(1)
        On Error Resume Next
        Dim vbc As VBComponent
        Set vbc = .VBComponents(module) 'Microsoft Visual Basic for Applications Extensibility x.x
        If Err.Number = 0 Then .VBComponents.Remove vbc
        On Error GoTo 0
        .VBComponents.Import path
    End With
End Sub

Private Sub ImportObject(ByVal objectType As AcObjectType, ByVal path As String)
    Dim all As AllObjects
    Select Case objectType
    Case acForm
        Set all = CurrentProject.AllForms
    Case acReport
        Set all = CurrentProject.AllReports
    Case Else
        Err.Raise 93 'Invalid pattern string (https://support.microsoft.com/kb/146864)
    End Select
    Dim ao As AccessObject
    For Each ao In all
        DoCmd.Close objectType, ao.name, acSaveNo
    Next
    With New FileSystemObject
        Dim module As String
        module = .GetBaseName(path)
    End With
    With VBE.VBProjects(1)
        On Error Resume Next
        DoCmd.DeleteObject objectType, module
        On Error GoTo 0
        Application.LoadFromText objectType, module, path
    End With
End Sub

Private Function ImportSqlServer()
    Const CONNECT As String = "ODBC;DRIVER=SQL Server;SERVER=server1,port1;DATABASE=database1;UID=uid1;PWD=pwd1"
    Const TABLES As String = "pm1tables.csv"
    Const REMOTE_NAME As Long = 0
    Const LOCAL_NAME As Long = 1
    DeleteTables
    Dim table As Variant
    table = ReadCSV(CurrentProject.path & "\" & TABLES)
    Dim iTable As Long
    For iTable = 0 To UBound(table)
        DoCmd.TransferDatabase acImport, "ODBC Database", CONNECT, acTable, table(iTable, REMOTE_NAME), table(iTable, LOCAL_NAME)
    Next
    Application.Quit
End Function