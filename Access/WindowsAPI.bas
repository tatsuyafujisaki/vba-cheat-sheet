Option Compare Database
Option Explicit

' https://msdn.microsoft.com/library/windows/desktop/ms632673.aspx
Private Declare PtrSafe Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long

' https://msdn.microsoft.com/library/windows/desktop/ms633499.aspx
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Long

Sub BringToTop(ByVal windowTitle As String)
    Dim hWnd As Long: hWnd = FindWindow(vbEmpty, windowTitle)
    If hWnd <> 0 Then BringWindowToTop hWnd
End Sub
