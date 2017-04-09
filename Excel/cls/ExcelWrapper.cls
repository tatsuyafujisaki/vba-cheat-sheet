Option Explicit

Public Excel As Excel.Application

Private Sub Class_Initialize()
    Set Excel = New Excel.Application 'Microsoft Excel x.x Object Library
    With Excel
        .SheetsInNewWorkbook = 1
        .ScreenUpdating = False
        .DisplayAlerts = False
        .Workbooks.Add
        .Calculation = xlCalculationManual
    End With
End Sub

Private Sub Class_Terminate()
    With Excel
        .Visible = True
        .DisplayAlerts = True
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
    End With
    Set Excel = Nothing
End Sub