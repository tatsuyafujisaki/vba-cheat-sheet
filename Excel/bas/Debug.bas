Option Explicit

Private Sub RemoveCommentsForExcel()
    Dim vbc As VBComponent
    For Each vbc In ThisWorkbook.VBProject.VBComponents
        RemoveComments vbc.CodeModule
    Next
End Sub

Private Sub RemoveCommentsForAccess()
    Dim vbp As VBProject
    For Each vbp In Application.VBE.VBProjects
        Dim vbc As VBComponent
        For Each vbc In vbp.VBComponents
            RemoveComments vbc.CodeModule
        Next
    Next
End Sub

' Referenes "Microsoft Visual Basic for Applications Extensibility 5.x" for CodeModule
Private Sub RemoveComments(cm As CodeModule)
    Const nonComment As Long = -1
    Const blank As Long = -2
    Const fullComment As Long = -3
    Const doubleQuote As Long = 34
    Const singleQuote As Long = 39
    ReDim lineTypes(1 To cm.CountOfLines) As Long
    Dim iLine As Long
    For iLine = 1 To cm.CountOfLines
        lineTypes(iLine) = nonComment
        Dim line As String
        line = cm.Lines(iLine, 1)
        If Trim(line) = "" Then
            lineTypes(iLine) = blank
        ElseIf Left(Trim(line), 1) = Chr(singleQuote) Or IsContinualComment(cm, lineTypes, iLine) Then
            lineTypes(iLine) = fullComment
        Else
            Dim inDoubleQuotes As Boolean
            inDoubleQuotes = False
            Dim iPos As Long
            For iPos = 1 To Len(line)
                Dim c As String
                c = Mid(line, iPos, 1)
                If c = Chr(doubleQuote) Then
                    inDoubleQuotes = Not inDoubleQuotes
                ElseIf (c = Chr(singleQuote)) And (Not inDoubleQuotes) Then
                    lineTypes(iLine) = iPos
                    Exit For
                End If
            Next
        End If
    Next
    For iLine = cm.CountOfLines To 1 Step -1
        line = cm.Lines(iLine, 1)
        Select Case lineTypes(iLine)
            Case blank
                cm.DeleteLines iLine
            Case fullComment
                cm.DeleteLines iLine
            Case nonComment
            Case Else
                cm.ReplaceLine iLine, Mid(line, 1, lineTypes(iLine) - 1)
        End Select
    Next
    Erase lineTypes
End Sub

Private Function IsContinualComment(cm As CodeModule, lineTypes() As Long, ByVal iLine As Long) As Boolean
    Const fullComment As Long = -3
    If iLine = 1 Then 'IIf makes an error
        IsContinualComment = False
    Else
        IsContinualComment = ((lineTypes(iLine - 1) = fullComment Or 0 < lineTypes(iLine - 1)) And Right(cm.Lines(iLine - 1, 1), 2) = " _")
    End If
End Function

Private Sub CreateInsertsForAccess(ws As Worksheet)
    Dim table
    table = ws.UsedRange

    Dim nCols As Long
    nCols = UBound(table, 2)

    Dim iRow As Long
    For iRow = 1 To UBound(table)
        Dim sql As String
        sql = "CurrentDb.Execute " & Chr(34) & "INSERT INTO " & ws.Name & " VALUES(" & Quote(table(iRow, 1))
        Dim iCol As Long
        For iCol = 2 To nCols
            sql = sql & Chr(44) & Quote(table(iRow, iCol))
        Next
        sql = sql & ")" & Chr(34)
        Debug.Print sql
    Next
End Sub

Private Function Quote(ByVal s As String) As String
    Quote = Chr(39) & s & Chr(39)
End Function

Private Sub PrintVBComponents()
    Dim vbcs As Object
    Set vbcs = CreateObject("System.Collections.ArrayList")

    Dim vbc
    For Each vbc In ThisWorkbook.VBProject.VBComponents
        vbcs.Add vbc.Name 'vbc is VBComponent here
    Next
    vbcs.Sort
    For Each vbc In vbcs
        Debug.Print vbc 'vbc is String here
    Next
End Sub

Private Sub PrintPivotTables(ws As Worksheet)
    Dim pt As PivotTable
    For Each pt In ws.PivotTables
        Debug.Print pt.Name
    Next
End Sub

Private Sub PrintCollection(c As Collection)
    Dim item
    For Each item In c
        Debug.Print item
    Next
End Sub

Private Sub PrintButtonInfo(sht As Worksheet, ByVal buttonName As String)
    With sht.Buttons(buttonName)
        Debug.Print .Name
        Debug.Print .Caption
        Debug.Print .OnAction
        Debug.Print .Width
        Debug.Print .Height
        Debug.Print .Left
        Debug.Print .Top
    End With
End Sub

Private Sub Backup()
    Dim backupDir As String
    backupDir = ActiveWorkbook.Path & "\" & "backup"

    If Dir(backupDir, vbDirectory) = "" Then MkDir backupDir

    With New FileSystemObject
        Dim message As String
        message = InputBox("Please enter the commit message.")
        Shell "cmd /c echo f | xcopy /f /y " & Chr(34) & ActiveWorkbook.FullName & Chr(34) & " " & Chr(34) & ActiveWorkbook.Path & "\backup\" & .GetBaseName(ActiveWorkbook.FullName) & "_" & Format(Now, "yyyymmdd_hhMM") & IIf(Len(message), "_" & Replace(message, " ", "_"), "") & "." & .GetExtensionName(ActiveWorkbook.FullName) & Chr(34)
    End With
End Sub