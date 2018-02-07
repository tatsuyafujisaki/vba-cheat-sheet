Option Explicit

Private screenUpdating As Boolean
Private displayStatusBar As Boolean
Private enableEvent As Boolean
Private calculation As Long

Private Sub Class_Initialize()
    With Application
        screenUpdating = .ScreenUpdating
        .ScreenUpdating = False
        
        displayStatusBar = .displayStatusBar
        .displayStatusBar = False
        
        enableEvent = .EnableEvents
        .EnableEvents = False

        calculation = .Calculation
        .Calculation = xlCalculationManual
    End With
End Sub

Private Sub Class_Terminate()
    With Application
        .Calculation = calculation
        .EnableEvents = enableEvent
        .DisplayStatusBar = displayStatusBar
        .ScreenUpdating = screenUpdating
    End With
End Sub

' Usage
' Dim o As Optimizer
' Set o = New Optimizer

' The following syntax does not trigger Class_Initialize()
' Dim o As New Optimizer
