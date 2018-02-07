Option Explicit

Private Sub Main()
    ProcedureThatNeedsManualCalculation
End Sub

Private Sub ProcedureThatNeedsManualCalculation()
    Dim o As Optimizer
    Set o = New Optimizer

    ' Calculation is manual here.

    ProcedureThatNeedsAutomaticCalculation

    ' Calculation is manual here.
End Sub

Private Sub ProcedureThatNeedsAutomaticCalculation()
    Dim u As Unoptimizer
    Set u = New Unoptimizer

    ' Calculation is automatic here.
End Sub
