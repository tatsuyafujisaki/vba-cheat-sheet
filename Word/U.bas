Option Explicit

Private Function IsDocument(ByVal path As String) As Boolean
    With New FileSystemObject
        Dim ext As String
        ext = .GetExtensionName(path)
    End With
    IsDocument = (StrComp(ext, "docx", vbTextCompare) = 0) Or (StrComp(ext, "doc", vbTextCompare) = 0)
End Function

Private Function IsBackup(ByVal path As String) As Boolean
    With New FileSystemObject
        IsBackup = Left$(.GetBaseName(path), 1) = "~"
    End With
End Function

Private Sub MkDirIfNotExist(ByVal path As String)
    If Dir(path, vbDirectory) = vbNullString Then MkDir path
End Sub

Private Sub SaveUnsavedDocument(ByVal path As String)
    Dim d As document
    For Each d In Application.Documents
        If d.path = vbNullString Then
            With New FileSystemObject
                MkDirIfNotExist .GetParentFolderName(path)
            End With
            d.SaveAs path, wdFormatXMLDocument
            Exit Sub
        End If
    Next
End Sub

Private Sub CompareDocuments()
    Const InputDir As String = "C:\Input"
    Const OutputDir As String = "C:\Ouput"

    Dim fso As New FileSystemObject
    Dim dealDir As folder
    For Each dealDir In fso.GetFolder(InputDir).SubFolders
        Dim oldDir As String
        oldDir = fso.BuildPath(dealDir.path, "Old")

        Dim newDir As String
        newDir = fso.BuildPath(dealDir.path, "New")

        If Dir(oldDir, vbDirectory) = vbNullString Then
            Debug.Print oldDir & " does not exist"
        ElseIf Dir(newDir, vbDirectory) = vbNullString Then
            Debug.Print newDir & " does not exist"
        Else
            Dim oldOtDirs As Folders
            Set oldOtDirs = fso.GetFolder(oldDir).SubFolders

            Dim newOtDirs As Folders
            Set newOtDirs = fso.GetFolder(newDir).SubFolders

            If oldOtDirs.Count <> newOtDirs.Count Then
                Debug.Print "# of outTypeDirs differ at " & dealDir
            Else
                Dim otDir As folder
                For Each otDir In oldOtDirs
                    Dim file As file
                    Dim fileCount As Long
                    fileCount = 1
                    For Each file In otDir.Files

                        Dim oldFile As String
                        oldFile = file.path

                        Dim newFile As String
                        newFile = Replace(file.path, "\Old\", "\New\")

                        If IsDocument(oldFile) And Not IsBackup(oldFile) Then

                            Dim oldDoc As document
                            Set oldDoc = Documents.Open(oldFile)

                            Dim newDoc As document
                            Set newDoc = Documents.Open(newFile)

                            Application.CompareDocuments oldDoc, newDoc
                            SaveUnsavedDocument fso.BuildPath(OutputDir, dealDir.Name & "_" & otDir.Name & "_" & fileCount & ".docx")
                            CloseAllDocumentsButMe oldDoc
                            CloseAllDocumentsButMe newDoc
                            fileCount = fileCount + 1
                        End If
                    Next
                Next
            End If
        End If
    Next
End Sub

Private Sub CloseAllDocumentsButMe(me1 As document)
    Dim d As document
    For Each d In Application.Documents
        If d.FullName <> me1.FullName Then d.Close wdDoNotSaveChanges
    Next
End Sub
