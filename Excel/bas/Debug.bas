Option Explicit

Sub RemoveCommentsForExcel()
    Dim vbc As VBComponent
    For Each vbc In ThisWorkbook.VBProject.VBComponents
        RemoveComments vbc.CodeModule
    Next
End Sub

Sub RemoveCommentsForAccess()
    Dim vbp As VBProject
    For Each vbp In Application.VBE.VBProjects
        Dim vbc As VBComponent
        For Each vbc In vbp.VBComponents
            RemoveComments vbc.CodeModule
        Next
    Next
End Sub

Sub RemoveComments(ByVal cm As CodeModule) ' Microsoft Visual Basic for Applications Extensibility x.x
    Const nonComment As Long = -1
    Const blank As Long = -2
    Const fullComment As Long = -3
    Const doubleQuote As Long = 34
    Const singleQuote As Long = 39
    ReDim lineTypes(1 To cm.CountOfLines) As Long
    Dim lineIndex As Long
    For lineIndex = 1 To cm.CountOfLines
        lineTypes(lineIndex) = nonComment
        Dim line As String
        line = cm.Lines(lineIndex, 1)
        If Trim$(line) = vbNullString Then
            lineTypes(lineIndex) = blank
        ElseIf Left$(Trim$(line), 1) = Chr$(singleQuote) Or IsContinualComment(cm, lineTypes, lineIndex) Then
            lineTypes(lineIndex) = fullComment
        Else
            Dim inDoubleQuotes As Boolean
            inDoubleQuotes = False
            Dim positionIndex As Long
            For positionIndex = 1 To Len(line)
                Dim c As String
                c = Mid$(line, positionIndex, 1)
                If c = Chr$(doubleQuote) Then
                    inDoubleQuotes = Not inDoubleQuotes
                ElseIf (c = Chr$(singleQuote)) And (Not inDoubleQuotes) Then
                    lineTypes(lineIndex) = positionIndex
                    Exit For
                End If
            Next
        End If
    Next
    For lineIndex = cm.CountOfLines To 1 Step -1
        line = cm.Lines(lineIndex, 1)
        Select Case lineTypes(lineIndex)
            Case blank
                cm.DeleteLines lineIndex
            Case fullComment
                cm.DeleteLines lineIndex
            Case nonComment
            Case Else
                cm.ReplaceLine lineIndex, Mid$(line, 1, lineTypes(lineIndex) - 1)
        End Select
    Next
    Erase lineTypes
End Sub

Function IsContinualComment(ByVal cm As CodeModule, ByRef lineTypes() As Long, ByVal lineIndex As Long) As Boolean
    Const fullComment As Long = -3
    If lineIndex = 1 Then 'IIf makes an error
        IsContinualComment = False
    Else
        IsContinualComment = ((lineTypes(lineIndex - 1) = fullComment Or 0 < lineTypes(lineIndex - 1)) And Right$(cm.Lines(lineIndex - 1, 1), 2) = " _")
    End If
End Function

Sub CreateInsertsForAccess(ByVal ws As Worksheet)
    Dim table As Variant
    table = ws.UsedRange

    Dim columnCount As Long
    columnCount = UBound(table, 2)

    Dim rowIndex As Long
    For rowIndex = 1 To UBound(table)
        Dim sql As String
        sql = "CurrentDb.Execute " & Chr$(34) & "INSERT INTO " & ws.name & " VALUES(" & Quote(table(rowIndex, 1))
        Dim columnIndex As Long
        For columnIndex = 2 To columnCount
            sql = sql & Chr$(44) & Quote(table(rowIndex, columnIndex))
        Next
        sql = sql & ")" & Chr$(34)
        Debug.Print sql
    Next
End Sub

Function Quote(ByVal s As String) As String
    Quote = Chr$(39) & s & Chr$(39)
End Function

Sub PrintVBComponents()
    Dim vbcs As Object
    Set vbcs = CreateObject("System.Collections.ArrayList")

    Dim vbc As Variant
    For Each vbc In ThisWorkbook.VBProject.VBComponents
        vbcs.Add vbc.name 'vbc is VBComponent here
    Next
    vbcs.Sort
    For Each vbc In vbcs
        Debug.Print vbc 'vbc is String here
    Next
End Sub

Sub PrintPivotTables(ByVal ws As Worksheet)
    Dim pt As PivotTable
    For Each pt In ws.PivotTables
        Debug.Print pt.name
    Next
End Sub

Sub PrintCollection(ByVal c As Collection)
    Dim item As Variant
    For Each item In c
        Debug.Print item
    Next
End Sub

Sub PrintButtonInfo(ByVal sht As Worksheet, ByVal buttonName As String)
    With sht.Buttons(buttonName)
        Debug.Print .name
        Debug.Print .Caption
        Debug.Print .OnAction
        Debug.Print .Width
        Debug.Print .Height
        Debug.Print .Left
        Debug.Print .Top
    End With
End Sub

Sub PrintAddins()
    Dim addin As addin

    For Each addin In Application.AddIns2
        Debug.Print addin.Name
        Debug.Print addin.Path
        Debug.Print addin.FullName
        Debug.Print addin.Installed
        Debug.Print
    Next
End Sub

Sub Backup()
    Dim backupDir As String
    backupDir = ActiveWorkbook.PATH & "\" & "backup"

    If Dir(backupDir, vbDirectory) = vbNullString Then MkDir backupDir

    With New FileSystemObject
        Dim message As String
        message = InputBox("Please enter the commit message.")
        Shell "cmd /c echo f | xcopy /f /y " & Chr$(34) & ActiveWorkbook.FullName & Chr$(34) & " " & Chr$(34) & ActiveWorkbook.PATH & "\backup\" & .GetBaseName(ActiveWorkbook.FullName) & "_" & Format$(Now, "yyyymmdd_hhMM") & IIf(Len(message), "_" & Replace(message, " ", "_"), vbNullString) & "." & .GetExtensionName(ActiveWorkbook.FullName) & Chr$(34)
    End With
End Sub