Private Function FileExists(ByVal filePath As String) As Boolean
    FileExists = Dir$(filePath) <> vbNullString
End Function

Private Function DirectoryExists(ByVal directoryPath As String) As Boolean
    DirectoryExists = Dir$(directoryPath, vbDirectory) <> vbNullString
End Function