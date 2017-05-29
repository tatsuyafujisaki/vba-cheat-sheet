Option Explicit

' Note
' LeftIndent is ignored when CharacterUnitLeftIndent is non-zero (positive or negative), or CharacterUnitFirstLineIndent is negative.
' FirstLineIndent is ignored when CharacterUnitFirstLineIndent is non-zero (positive or negative).

' ParagraphFormat.CharacterUnitLeftIndent
' https://msdn.microsoft.com/library/office/ff836968.aspx

' ParagraphFormat.CharacterUnitFirstLineIndent
' https://msdn.microsoft.com/library/office/ff840585.aspx

' ParagraphFormat.LeftIndent
' https://msdn.microsoft.com/library/office/ff837464.aspx

' ParagraphFormat.FirstLineIndent
' https://msdn.microsoft.com/library/office/ff836045.aspx

Private Sub Demo()
    Dim p As Paragraph
    Set p = ThisDocument.Paragraphs.First
    
    ResetIndent p
End Sub

Private Sub ResetIndent(ByRef p As Paragraph)
    p.Range.Select
    
    With Selection.ParagraphFormat
        .CharacterUnitLeftIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LeftIndent = 0
        .firstLineIndent = 0
    End With
    
    Selection.HomeKey wdStory
End Sub

Public Sub SetHangingIndentInCharacterUnit(ByRef p As Paragraph, ByVal indent As Single, ByVal additionalIndentForHanging As Single)
    p.Range.Select
    
    With Selection.ParagraphFormat
        .LeftIndent = 0
        .firstLineIndent = 0
        .CharacterUnitLeftIndent = indent
        .CharacterUnitFirstLineIndent = -additionalIndentForHanging
        
        Dim hangingIndent As Single
        hangingIndent = (.CharacterUnitFirstLineIndent + .CharacterUnitLeftIndent) * -1

        Debug.Print "First-line indent:" & .CharacterUnitLeftIndent & " characters"
        Debug.Print "Hanging indent:" & hangingIndent & " characters"
        Debug.Print "Second-line indent (First-line indent + hanging indent):" & .CharacterUnitLeftIndent + hangingIndent & " characters"
    End With
    
    Selection.HomeKey wdStory
End Sub


Public Sub SetHangingIndentInPoint(ByRef p As Paragraph, ByVal indent As Single, ByVal additionalIndentForHanging As Single)
    p.Range.Select
    
    With Selection.ParagraphFormat
        .CharacterUnitLeftIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LeftIndent = indent + additionalIndentForHanging
        .firstLineIndent = -additionalIndentForHanging
        
        Dim firstLineIndent As Single
        firstLineIndent = .LeftIndent + .firstLineIndent
        
        Dim hangingIndent As Single
        hangingIndent = .firstLineIndent * -1
        
        Debug.Print "First-line indent:" & firstLineIndent & " points"
        Debug.Print "Hanging indent:" & hangingIndent & " points"
        Debug.Print "Second-line indent (First-line indent + hanging indent):" & firstLineIndent + hangingIndent & " points"
    End With

    Selection.HomeKey wdStory
End Sub
