Option Compare Database
Option Explicit

Sub TransferObjects()
    Dim ao As AccessObject
    For Each ao In CurrentProject.AllForms
        TransferObject acForm, ao.name
    Next
    For Each ao In CurrentProject.AllMacros
        TransferObject acMacro, ao.name
    Next
    For Each ao In CurrentProject.AllModules
        TransferObject acModule, ao.name
    Next
    For Each ao In CurrentProject.AllReports
        TransferObject acReport, ao.name
    Next
    Dim td As DAO.TableDef
    For Each td In CurrentDb.TableDefs
        If Left$(td.name, 4) <> "MSys" Then TransferObject acTable, td.name
    Next
    Dim qd As DAO.QueryDef
    For Each qd In CurrentDb.QueryDefs
        If Left$(qd.name, 1) <> "~" Then TransferObject acQuery, qd.name
    Next
    MsgBox "Done!"
End Sub

Sub TransferObject(ByVal objectType As AcObjectType, ByVal name As String)
    Dim exportTo As String: exportTo = CurrentProject.path & "\" & "Database1.accdb"
    DoCmd.TransferDatabase acExport, "Microsoft Access", exportTo, objectType, name, name
End Sub
