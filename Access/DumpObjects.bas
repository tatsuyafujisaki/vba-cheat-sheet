Option Compare Database
Option Explicit

Private Sub DumpObjects()
    On Error Resume Next

    Dim outRoot As String
    outRoot = GetBuiltPath(CurrentProject.path, "dump")
    MkDirIfNotExist outRoot

    Dim outDir As String
    outDir = GetBuiltPath(outRoot, "macros")
    MkDirIfNotExist outDir

    Dim ao As AccessObject
    For Each ao In CurrentProject.AllMacros
        SaveAsText acMacro, ao.name, GetBuiltPath(outDir, ao.name)
    Next

    outDir = GetBuiltPath(outRoot, "modules")
    MkDirIfNotExist outDir

    For Each ao In CurrentProject.AllModules
        SaveAsText acModule, ao.name, GetBuiltPath(outDir, ao.name)
    Next

    outDir = GetBuiltPath(outRoot, "forms")
    MkDirIfNotExist outDir

    For Each ao In CurrentProject.AllForms
        SaveAsText acForm, ao.name, GetBuiltPath(outDir, ao.name)
    Next

    outDir = GetBuiltPath(outRoot, "reports")
    MkDirIfNotExist outDir

    For Each ao In CurrentProject.AllReports
        SaveAsText acReport, ao.name, GetBuiltPath(outDir, ao.name)
    Next

    outDir = GetBuiltPath(outRoot, "tables")
    MkDirIfNotExist outDir

    Dim td As DAO.TableDef
    For Each td In CurrentDb.TableDefs
        If Left(td.name, 4) <> "MSys" Then ExportXML acExportTable, td.name, GetBuiltPath(outDir, td.name & ".xml"), GetBuiltPath(outDir, td.name & ".xsd")
    Next

    outDir = GetBuiltPath(outRoot, "queries")
    MkDirIfNotExist outDir

    Dim qd As DAO.QueryDef
    For Each qd In CurrentDb.QueryDefs
        If Left(qd.name, 1) <> "~" Then SaveAsText acQuery, qd.name, GetBuiltPath(outDir, qd.name)
    Next

    MsgBox "Done!"
    On Error GoTo 0
End Sub
