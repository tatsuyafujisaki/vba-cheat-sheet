Option Explicit

Private screenUpdating As Boolean
Private displayStatusBar As Boolean
Private enableEvent As Boolean
Private caluclation As Long

Private Sub Class_Initialize()
    With Application
        screenUpdating = .screenUpdating
        .screenUpdating = False
        
        displayStatusBar = Application.displayStatusBar
        .displayStatusBar = False
        
        enableEvent = Application.EnableEvents
        .EnableEvents = False

        caluclation = .Calculation
        .Calculation = xlCalculationManual
    End With
End Sub

Private Sub Class_Terminate()
    With Application
        .Calculation = caluclation
        .EnableEvents = enableEvent
        .displayStatusBar = displayStatusBar
        .screenUpdating = screenUpdating
    End With
End Sub

' Usage
' Dim o As Optimizer
' Set o = New Optimizer

' The following does not trigger Class_Initialize().
' Dim o As New Optimizer
