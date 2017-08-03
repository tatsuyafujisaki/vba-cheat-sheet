Option Explicit

' Here are functions that use Me (ThisDocument).

Private Sub Document_Open()
    Me.ActiveWindow.View.SplitSpecial = wdPaneRevisions
End Sub

Private Sub MarkShapes()
    Dim s As Shape

    Me.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekCurrentPageFooter

    For Each s In Me.Shapes
        s.Select
        Selection.TypeText s.Name
    Next
End Sub

Private Sub CloseAllDocuments()
    Dim d As document
    For Each d In Application.Documents
        If d.FullName <> Me.FullName Then d.Close wdDoNotSaveChanges
    Next
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
