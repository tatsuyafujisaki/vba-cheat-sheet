Option Explicit

Private Function TableExists(ByVal name As String) As Boolean
    TableExists = False
    Dim td As DAO.TableDef
    For Each td In CurrentDb.TableDefs
        If td.name = name Then
            TableExists = True
            Exit For
        End If
    Next
End Function

Private Function QueryExists(ByVal name As String) As Boolean
    QueryExists = False
    Dim qd As DAO.QueryDef
    For Each qd In CurrentDb.QueryDefs
        If qd.name = name Then
            QueryExists = True
            Exit For
        End If
    Next
End Function

Private Function FormExists(ByVal name As String) As Boolean
    FormExists = False
    Dim ao As AccessObject
    For Each ao In CurrentProject.AllForms
        If ao.name = name Then
            FormExists = True
            Exit For
        End If
    Next
End Function

Private Function ReportExists(ByVal name As String) As Boolean
    ReportExists = False
    Dim ao As AccessObject
    For Each ao In CurrentProject.AllReports
        If ao.name = name Then
            ReportExists = True
            Exit For
        End If
    Next
End Function

Private Function MacroExists(ByVal name As String) As Boolean
    MacroExists = False
    Dim ao As AccessObject
    For Each ao In CurrentProject.AllMacros
        If ao.name = name Then
            MacroExists = True
            Exit For
        End If
    Next
End Function

Private Function ModuleExists(ByVal name As String) As Boolean
    ModuleExists = False
    Dim ao As AccessObject
    For Each ao In CurrentProject.AllModules
        If ao.name = name Then
            ModuleExists = True
            Exit For
        End If
    Next
End Function