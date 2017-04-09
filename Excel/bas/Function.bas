Option Explicit

Private Function Printf(ByVal format As String, ParamArray args()) As String
    Dim i As Long
    For i = 0 To UBound(args)
        format = Replace(format, "{" & i & "}", args(i))
    Next
    Printf = format
End Function

Private Function IsOdd(ByVal x As Long) As Boolean
    IsOdd = x Mod 2
End Function

Private Function IsEven(ByVal x As Long) As Boolean
    IsEven = x Mod 2 = 0
End Function

Private Function AreNumeric(arr) As Boolean
    AreNumeric = True
    Dim v
    For Each v In arr
        If Not IsNumeric(v) Then
            AreNumeric = False
            Exit For
        End If
    Next
End Function

Private Function GetMedian(arr)
    Dim al As Object
    Set al = CreateObject("System.Collections.ArrayList")

    Dim n
    For Each n In arr
        al.Add CDbl(n)
    Next

    al.Sort 'log(n*Log(n))

    Dim mid As Long
    mid = Fix(al.Count / 2)

    GetMedian = IIf(al.Count Mod 2, al(mid), (al(mid) + al(mid + 1)) / 2)
End Function

Private Function GetBuiltPath(ByVal folder As String, ByVal file As String) As String
    With New FileSystemObject
        GetBuiltPath = .BuildPath(folder, file)
    End With
End Function

Private Function GetParentFolder(ByVal folder As String) As String
    With New FileSystemObject
        GetParentFolder = .GetParentFolderName(folder)
    End With
End Function

'Usage: GetFiles(New Collection, "C:\foo", Array("txt", "sql"))
Private Function GetFiles(files As Collection, ByVal folder As String, Optional extensions = Null) As Collection
    With New FileSystemObject
        Dim subfolder As folder
        For Each subfolder In .GetFolder(folder).SubFolders
            Set files = GetFiles(files, subfolder, extensions)
        Next
        If IsNull(extensions) Then
            Dim file As file
            For Each file In .GetFolder(folder).files
                files.Add file.path
            Next
        Else
            For Each file In .GetFolder(folder).files
                If InArray(extensions, .GetExtensionName(file)) Then files.Add file.path
            Next
        End If
    End With
    Set GetFiles = files
End Function

Private Function InArray(xs, ByVal findMe As String) As Boolean
    InArray = False
    Dim element
    For Each element In xs
        If StrComp(element, findMe, vbTextCompare) = 0 Then
            InArray = True
            Exit For
        End If
    Next
End Function

Private Function IsAbsolutePath(ByVal path As String) As Boolean
    With New FileSystemObject 'Microsoft Scripting Runtime
        IsAbsolutePath = StrComp(path, .GetAbsolutePathName(path), vbTextCompare) = 0
    End With
End Function

Private Function GetAbsolutePath(ByVal path As String) As String
    GetAbsolutePath = IIf(IsAbsolutePath(path), path, ThisWorkbook.path & "\" & path) 'Use CurrentProject.Path for Access 'CurDir is not useful
End Function

Private Function IsProduction() As Boolean
    Const PRD_DIR As String = "path/to/production"
    IsProduction = ThisWorkbook.path = PRD_DIR
End Function

Private Function GetOSVersion() As String
    Select Case Application.OperatingSystem
    Case "Windows (32-bit) NT 6.01"
        GetOSVersion = "Windows 7"
    Case "Windows (32-bit) NT 5.01"
        GetOSVersion = "Windows XP"
    Case Else
        Err.Raise 93 'Invalid pattern string (https://support.microsoft.com/kb/146864)
    End Select
End Function

Private Function IsInvalid(v) As Boolean
    On Error Resume Next
    IsInvalid = (v = "")
    IsInvalid = IsInvalid And (Err.Number = 0)
    On Error GoTo 0
    If IsInvalid Then Exit Function
    On Error Resume Next
    IsInvalid = IsEmpty(v)
    IsInvalid = IsInvalid And (Err.Number = 0)
    On Error GoTo 0
    If IsInvalid Then Exit Function
    On Error Resume Next
    IsInvalid = IsNull(v)
    IsInvalid = IsInvalid And (Err.Number = 0)
    On Error GoTo 0
    If IsInvalid Then Exit Function
    On Error Resume Next
    IsInvalid = (v Is Nothing)
    IsInvalid = IsInvalid And (Err.Number = 0)
    On Error GoTo 0
End Function

Private Function Quote(ByVal s As String) As String
    Quote = Chr(39) & s & Chr(39)
End Function

Private Function DoubleQuote(ByVal s As String) As String
    DoubleQuote = Chr(34) & s & Chr(34)
End Function

Private Function QuoteOrNull(ByVal s) As String
    QuoteOrNull = IIf(IsNull(s), "Null", Chr(39) & s & Chr(39))
End Function

Private Function Bracket(ByVal s As String) As String
    Bracket = Chr(40) & s & Chr(41)
End Function

Private Function GetNumberSignedDate() As String
    GetNumberSignedDate = Chr(35) & Date & Chr(35)
End Function

Private Function GetGoldenRatio() As Double
    GetGoldenRatio = (1 + Sqr(5)) / 2
End Function

Private Function GetDesktopPath() As String
    With New WshShell 'Windows Script Host Object Model
        GetDesktopPath = .SpecialFolders("Desktop")
    End With
End Function

Private Function SheetExists(ByVal wsName As String) As Boolean
    On Error Resume Next
    SheetExists = Not (ThisWorkbook.Sheets(wsName) Is Nothing)
    On Error GoTo 0
End Function

Private Function GetNameDept()
    With New WshShell 'Windows Script Host Object Model
        With .Exec("net user /domain " & Environ("USERNAME"))
            Dim s As String
            s = .StdOut.ReadAll
        End With
    End With
    With New RegExp 'Microsoft VBScript Regular Expressions x.x
        .Pattern = "Full Name[ ]+([^\r]+)"

        Dim mc As MatchCollection
        Set mc = .Execute(s)

        Dim nd(1)
        nd(0) = mc(0).SubMatches(0)

        .Pattern = "Comment[ ]+([^\r]+)"
        Set mc = .Execute(s)
        nd(1) = Split(mc(0).SubMatches(0), "/")(1)
    End With
    GetNameDept = nd
End Function

Private Function GetFormattedDouble(ByVal d As Double) As String
    GetFormattedDouble = format(WorksheetFunction.RoundDown(d, 3), "0,000.000")
End Function

Private Function GetCSV(ParamArray pa()) As String
    GetCSV = Join(Array(pa), ",")
End Function

Private Function GetNonEmptyCellsInColumn(ws As Worksheet, ByVal colIndex As Long) As Range
    Dim col As Range
    Set col = Intersect(ws.UsedRange, ws.columns(colIndex)) 'CurrentRegion doesn't work with entire column

    Dim singleCell As Range
    For Each singleCell In col
        If Not IsEmpty(singleCell) Then
            If GetNonEmptyCellsInColumn Is Nothing Then 'IIf make an error
                Set GetNonEmptyCellsInColumn = singleCell
            Else
                Set GetNonEmptyCellsInColumn = Union(GetNonEmptyCellsInColumn, singleCell)
            End If
        End If
    Next
End Function

Private Function A1ToR1C1(ByVal address As String) As String
    A1ToR1C1 = Application.ConvertFormula(address, xlA1, xlR1C1, xlAbsolute, Sheets(1).Cells(1, 1))
End Function

Private Function R1C1ToA1(ByVal address As String) As String
    R1C1ToA1 = Application.ConvertFormula(address, xlR1C1, xlA1, xlAbsolute, Sheets(1).Cells(1, 1))
End Function

Private Function IndexToAlphabet(ByVal columnIndex As Long) As String
    IndexToAlphabet = Replace(Cells(1, columnIndex).address(False, False), "1", "")
End Function

Private Function AlphabetToIndex(ByVal colAlphabet As String) As Long
    AlphabetToIndex = columns(colAlphabet).Column
End Function

Private Function AddWorksheet(wb As Workbook, ByVal codeName As String, ByVal name As String, Optional ByVal before As String = Empty, Optional ByVal after As String = Empty) As Worksheet
    Dim sh As Worksheet
    If before <> "" Then
        Set sh = wb.Sheets.Add(before:=wb.Sheets(before))
    ElseIf after <> "" Then
        Set sh = wb.Sheets.Add(after:=wb.Sheets(after))
    Else
        Set sh = wb.Sheets.Add
    End If
    wb.VBProject.VBComponents(sh.codeName).name = codeName
    sh.name = name
    Set AddWorksheet = sh
End Function
