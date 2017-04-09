Option Explicit

Private Sub AlignTop()
    ActiveWindow.Selection.ShapeRange.Align msoAlignTops, msoFalse
End Sub

Private Sub AlignBottom()
    ActiveWindow.Selection.ShapeRange.Align msoAlignBottoms, msoFalse
End Sub

Private Sub AlignLeft()
    ActiveWindow.Selection.ShapeRange.Align msoAlignLefts, msoFalse
End Sub

Private Sub AlignRight()
    ActiveWindow.Selection.ShapeRange.Align msoAlignRights, msoFalse
End Sub

Private Sub AlignCenterV()
    ActiveWindow.Selection.ShapeRange.Align msoAlignCenters, msoFalse
End Sub

Private Sub AlignCenterH()
    ActiveWindow.Selection.ShapeRange.Align msoAlignMiddles, msoFalse
End Sub

Private Sub DistributeHorizontally()
    ActiveWindow.Selection.ShapeRange.Distribute msoDistributeHorizontally, msoFalse
End Sub

Private Sub DistributeVertically()
    ActiveWindow.Selection.ShapeRange.Distribute msoDistributeVertically, msoFalse
End Sub

Private Sub SetShapeSize()
    Const HEIGHT As Long = 20
    Const WIDTH As Long = 100
    If ActiveWindow.Selection.Type = ppSelectionShapes Then
        Dim sr As Shape
        For Each sr In ActiveWindow.Selection.ShapeRange
            sr.WIDTH = WIDTH
            sr.HEIGHT = HEIGHT
        Next
    End If
End Sub

Private Sub SetFont()
    If ActiveWindow.Selection.Type = ppSelectionShapes Then
        Dim tr As TextRange
        For Each tr In ActiveWindow.Selection.TextRange
            With tr.Font
                .Name = "Meiryo UI"
                .Size = 10
            End With
        Next
    End If
End Sub
