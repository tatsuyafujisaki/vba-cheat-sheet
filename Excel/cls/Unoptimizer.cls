Option Explicit

Private screenUpdating As Boolean
Private displayStatusBar As Boolean
Private enableEvent As Boolean
Private calculation As Long

Private Sub Class_Initialize()
    With Application
        screenUpdating = .ScreenUpdating
        .ScreenUpdating = True
        
        displayStatusBar = .DisplayStatusBar
        .DisplayStatusBar = True
        
        enableEvent = .EnableEvents
        .EnableEvents = True

        calculation = .Calculation
        .Calculation = xlCalculationAutomatic
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
' Dim u As Unoptimizer
' Set u = New Unoptimizer

' The following syntax does not trigger Class_Initialize()
' Dim u As New Unoptimizer
