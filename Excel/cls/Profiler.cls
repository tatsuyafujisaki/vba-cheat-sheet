Option Explicit

Private t As Single

Private Sub Class_Initialize()
    t = Timer
End Sub

Private Sub Class_Terminate()
    Debug.Print (Timer - t) & " seconds"
End Sub

' Usage
' Dim p As Profiler
' Set p = New Profiler
