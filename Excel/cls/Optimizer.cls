Option Explicit

Private su As Boolean
Private c As Long

Private Sub Class_Initialize()
    With Application
        su = .ScreenUpdating
        c = .Calculation
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With
End Sub

Private Sub Class_Terminate()
    With Application
        .Calculation = c
        .ScreenUpdating = su
    End With
End Sub

' Usage
' Dim o As Optimizer
' Set o = New Optimizer

' The following does not trigger Class_Initialize().
' Dim o As New Optimizer