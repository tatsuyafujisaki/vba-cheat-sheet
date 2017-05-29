Option Compare Database
Option Explicit

Private Sub ImportVBComponentsAndObjects(ByVal paths As Variant)
    Dim fso As New FileSystemObject
    Dim path As Variant
    For Each path In paths
        Select Case fso.GetExtensionName(path)
            Case "bas"
                ImportVBComponent path
            Case "cls"
                ImportVBComponent path
            Case "form"
                ImportObject acForm, path
            Case "report"
                ImportObject acReport, path
            Case Else
                Err.Raise 93 'Invalid pattern string (https://support.microsoft.com/kb/146864)
        End Select
    Next
End Sub

Private Sub ImportVBComponent(ByVal path As String)
    With New FileSystemObject
        Dim module As String
        module = .GetBaseName(path)
    End With
    With VBE.VBProjects(1)
        On Error Resume Next
        Dim vbc As VBComponent ' Microsoft Visual Basic for Applications Extensibility 5.3
        Set vbc = .VBComponents(module)
        If Err.Number = 0 Then .VBComponents.Remove vbc
        On Error GoTo 0
        .VBComponents.Import path
    End With
End Sub

Private Sub ImportObject(ByVal objectType As AcObjectType, ByVal path As String)
    Dim all As AllObjects
    Select Case objectType
    Case acForm
        Set all = CurrentProject.AllForms
    Case acReport
        Set all = CurrentProject.AllReports
    Case Else
        Err.Raise 93 'Invalid pattern string (https://support.microsoft.com/kb/146864)
    End Select
    Dim ao As AccessObject
    For Each ao In all
        DoCmd.Close objectType, ao.Name, acSaveNo
    Next
    With New FileSystemObject
        Dim module As String
        module = .GetBaseName(path)
    End With
    With VBE.VBProjects(1)
        On Error Resume Next
        DoCmd.DeleteObject objectType, module
        On Error GoTo 0
        Application.LoadFromText objectType, module, path
    End With
End Sub
