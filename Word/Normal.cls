Option Explicit

Private Sub Document_Open()
    ActiveWindow.View.SplitSpecial = wdPaneRevisions
End Sub

Private Sub CompareDocuments()
    Const InputDir As String = "C:\Foo\Output"
    Const OutputDir As String = "C:\Foo\WordCompResults"

    CloseAllDocuments

    Dim fso As New FileSystemObject
    Dim dealDir As folder
    For Each dealDir In fso.GetFolder(InputDir).SubFolders
        Dim oldDir As String: oldDir = fso.BuildPath(dealDir.path, "Old")
        Dim newDir As String: newDir = fso.BuildPath(dealDir.path, "New")
        If Dir(oldDir, vbDirectory) = "" Then
            Debug.Print oldDir & " does not exist"
        ElseIf Dir(newDir, vbDirectory) = "" Then
            Debug.Print newDir & " does not exist"
        Else
            Dim oldOtDirs As Folders: Set oldOtDirs = fso.GetFolder(oldDir).SubFolders
            Dim newOtDirs As Folders: Set newOtDirs = fso.GetFolder(newDir).SubFolders
            If oldOtDirs.Count <> newOtDirs.Count Then
                Debug.Print "# of outTypeDirs differ at " & dealDir
            Else
                Dim otDir As folder
                For Each otDir In oldOtDirs
                    Dim file As file
                    Dim nFiles As Long: nFiles = 1
                    For Each file In otDir.Files
                        Dim oldFile As String: oldFile = file.path
                        Dim newFile As String: newFile = Replace(file.path, "\Old\", "\New\")
                        If IsDocument(oldFile) And Not IsBackup(oldFile) Then
                            Dim oldDoc As Document: Set oldDoc = Documents.Open(oldFile)
                            Dim newDoc As Document: Set newDoc = Documents.Open(newFile)
                            Application.CompareDocuments oldDoc, newDoc
                            SaveUnsavedDocument fso.BuildPath(OutputDir, dealDir.Name & "_" & otDir.Name & "_" & nFiles & ".docx")
                            CloseAllDocuments
                            CloseAllDocuments
                            nFiles = nFiles + 1
                        End If
                    Next
                Next
            End If
        End If
    Next
End Sub

Private Function IsDocument(ByVal path As String) As Boolean
    With New FileSystemObject
        Dim ext As String: ext = .GetExtensionName(path)
    End With
    IsDocument = (StrComp(ext, "docx", vbTextCompare) = 0) Or (StrComp(ext, "doc", vbTextCompare) = 0)
End Function

Private Function IsBackup(ByVal path As String) As Boolean
    With New FileSystemObject
        IsBackup = Left(.GetBaseName(path), 1) = "~"
    End With
End Function

Private Sub CloseAllDocuments()
    Dim d As Document
    For Each d In Application.Documents
        If d.FullName <> ThisDocument.FullName Then d.Close wdDoNotSaveChanges
    Next
End Sub

Private Sub SaveUnsavedDocument(ByVal path As String)
    Dim d As Document
    For Each d In Application.Documents
        If d.path = "" Then
            With New FileSystemObject
                MkDirIfNotExist .GetParentFolderName(path)
            End With
            d.SaveAs path, wdFormatXMLDocument
            Exit Sub
        End If
    Next
End Sub

Private Sub MkDirIfNotExist(ByVal path As String)
    If Dir(path, vbDirectory) = "" Then MkDir path
End Sub

Private Sub SetTabstopAtRightEnd()
    With Selection
        Dim width As Single
        With .PageSetup
            width = .PageWidth - (.LeftMargin + .RightMargin)
        End With
        With .Paragraphs.TabStops
            .ClearAll
            .Add width, wdAlignTabRight
        End With
    End With
End Sub