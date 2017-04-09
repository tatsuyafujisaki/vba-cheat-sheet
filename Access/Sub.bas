Option Explicit

Private Sub HideNavigationWindow()
    DoCmd.NavigateTo "acNavigationCategoryObjectType"
    DoCmd.RunCommand acCmdWindowHide
End Sub

Private Sub UnhideNavigationWindow()
    DoCmd.SelectObject acTable, , True
End Sub

Private Sub RunCommand(ByVal command As String)
    Shell "cmd /c " & command
End Sub

Private Sub RunCommandWithWindowMaximized(ByVal command As String)
    Shell "cmd /c start /max " & command
End Sub

Private Sub SaveWithoutPrompt()
    DoCmd.RunCommand acSaveNo
End Sub

Private Function GetThisModuleName() As String
    GetThisModuleName = VBE.ActiveCodePane.CodeModule.name
End Function

Private Sub DeleteModule(ByVal name As String)
    VBE.VBProjects(1).VBComponents.Remove name
End Sub