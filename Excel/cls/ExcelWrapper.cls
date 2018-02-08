Option Explicit

Public Excel As New Excel.Application 'Microsoft Excel x.x Object Library

Private Sub Class_Initialize()
    Set Excel = Excel.Application
    With Excel
        .SheetsInNewWorkbook = 1
        .ScreenUpdating = False
        .DisplayAlerts = False
        .EnableEvents = False
        .Workbooks.Add
        .Calculation = xlCalculationManual
    End With
End Sub

Private Sub Class_Terminate()
    With Excel
        .Workbooks.Close
        .Quit
    End With
End Sub
