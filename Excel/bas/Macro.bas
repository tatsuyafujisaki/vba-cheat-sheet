Option Explicit

Public Function Env(ByVal s As String) As String
    Env = Environ$(s)
End Function

Public Function GetCSV(ByVal r As Range) As String
    GetCSV = vbNullString
    Dim c As Range
    For Each c In r
        GetCSV = GetCSV & "," & c.Value
    Next
    GetCSV = Right$(GetCSV, Len(GetCSV) - 1)
End Function

Public Function IsFirstDayOfMonth(ByVal d As Date) As Boolean
    IsFirstDayOfMonth = Day(d) = 1
End Function

Public Function IsNewYearsDay(ByVal d As Date) As Boolean
    IsNewYearsDay = (Month(d) = 1) And (Day(d) = 1)
End Function

Public Function GetMondayOfSameWeek(ByVal d As Date) As Date
    GetMondayOfSameWeek = DateAdd("d", 2 - Weekday(d), d)
End Function

Public Function GetFridayOfSameWeek(ByVal d As Date) As Date
    GetFridayOfSameWeek = DateAdd("d", 6 - Weekday(d), d)
End Function

Public Function GetMondayOfLastWeek(ByVal d As Date) As Date
    GetMondayOfLastWeek = DateAdd("d", -Weekday(d) - 5, d)
End Function

Public Function GetFridayOfLastWeek(ByVal d As Date) As Date
    GetFridayOfLastWeek = DateAdd("d", -Weekday(d) - 1, d)
End Function

Public Function GetFirstDayOfLastMonth(ByVal d As Date) As Date
    GetFirstDayOfLastMonth = DateSerial(Year(d), Month(d) - 1, 1)
End Function
