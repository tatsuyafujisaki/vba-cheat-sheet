Option Explicit

Sub Demo()
    Const URL As String = "https://www.jasdec.com/reading/sonota.php?isin_4="
    Const ISIN As String = "JP582653BA93"
    Const KEY_TO_FIND As String = "<span class=""hy"">"

    With CreateObject("MSXML2.XMLHTTP.6.0")
        .Open "GET", URL & ISIN, False
        .Send
        Do While .readyState <> 4
            Application.Wait Now + TimeValue("0:00:01")
        Loop

        Dim html As String
        html = .responseText

    End With

    Debug.Print GetCharset(html)

    ' RegExp does not work if newline is contained and .MultiLine does not help
    html = Replace(html, KEY_TO_FIND & vbLf, KEY_TO_FIND)

    With New RegExp ' Microsoft VBScript Regular Expressions x.x
        .Pattern = KEY_TO_FIND & "(.+)</span>"
        .Global = True
        Dim mc As MatchCollection
        Set mc = .Execute(html)
        Debug.Print mc(3).SubMatches(0) ' LIBOR + alpha
    End With
End Sub

Function GetCharset(ByVal html As String) As String
    With New RegExp ' Microsoft VBScript Regular Expressions x.x
        .Pattern = "charset=(\w+)\W"
        Dim mc As MatchCollection
        Set mc = .Execute(html)
        GetCharset = mc(0).SubMatches(0)
    End With
End Function