Option Explicit

' Here are functions that use Me (ThisDocument).

Private Sub Document_Open()
    Me.ActiveWindow.View.SplitSpecial = wdPaneRevisions
End Sub

Private Sub MarkShapes()
    Dim shape As Shape
    Dim hf As HeaderFooter
    
    Me.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekCurrentPageFooter
    
    ' Counterintuitively, Sections(1) returns shapes in all sections.
    ' Counterintuitively, the following iteration through Headers returns shapes not only in the header but also in the footer too.
    For Each shape In Me.Sections(1).Headers(wdHeaderFooterPrimary).Shapes
        shape.Select
        Selection.TypeText "Shape in HeaderFooter > " & shape.Name
    Next
    
    For Each shape In Me.Shapes
        shape.Select
        Selection.TypeText "Shape in Body > " & shape.Name
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
