Option Explicit

' Microsoft Beefs Up VBScript with Regular Expressions
' https://msdn.microsoft.com/en-us/library/ms974570.aspx

Private Sub Demo()
    With New RegExp 'Microsoft VBScript Regular Expressions x.x
        .IgnoreCase = True 'Default is false
        'Test method
        .Pattern = "^\d+$"
        Debug.Print .test("01234") 'True
        Debug.Print .test("01A34") 'False

        'Replace method
        .Pattern = "[A-Za-z]+"
        Debug.Print .Replace("私はMikeです。", "マイク")

        'Global property
        .Pattern = "ABC"
        Debug.Print .Replace("ABCDEF ABCDEF ABCDEF", "abc") 'abcDEF ABCDEF ABCDEF
        .Global = True
        Debug.Print .Replace("ABCDEF ABCDEF ABCDEF", "abc") 'abcDEF abcDEF abcDEF

        'MultiLine property
        Dim s As String: s = "ABC" & vbCrLf & "DEF" & vbCrLf & "GHI" & vbCrLf
        .Pattern = "^D"
        Debug.Print .test(s) 'False
        .MultiLine = True
        Debug.Print .test(s) 'True

        'Execute method
        Dim mc As MatchCollection
        Dim m As Match
        .Pattern = "[A-Z]+"
        .Global = True
        Set mc = .Execute("ABC DEFG HIJKL MNOPQR STUVWXY")
        Debug.Print "mc.Count = " & mc.Count 'mc.Count = 5
        Dim i As Long
        For i = 0 To mc.Count - 1
            Set m = mc.Item(i)
            Debug.Print "FirstIndex = " & m.FirstIndex & " Length = " & m.Length & " Value = " & m.Value
        Next

        'SubMatches method
        .Pattern = "([A-Z])([a-z]+)"
        .Global = True
        Set mc = .Execute("Book Pen Apple Flower Sea Tree")
        Debug.Print "mc.Count = " & mc.Count 'mc.Count = 6
        For i = 0 To mc.Count - 1
            Set m = mc(i)
            Debug.Print "FirstIndex = " & m.FirstIndex & " Length = " & m.Length & " Value = " & m.Value & " SubMatches(0) = " & m.SubMatches(0) & " SubMatches(1) = " & m.SubMatches(1)
        Next
    End With
End Sub