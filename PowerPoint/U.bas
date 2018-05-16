Option Explicit

Sub AlignTop()
    ActiveWindow.Selection.ShapeRange.Align msoAlignTops, msoFalse
End Sub

Sub AlignBottom()
    ActiveWindow.Selection.ShapeRange.Align msoAlignBottoms, msoFalse
End Sub

Sub AlignLeft()
    ActiveWindow.Selection.ShapeRange.Align msoAlignLefts, msoFalse
End Sub

Sub AlignRight()
    ActiveWindow.Selection.ShapeRange.Align msoAlignRights, msoFalse
End Sub

Sub AlignCenterV()
    ActiveWindow.Selection.ShapeRange.Align msoAlignCenters, msoFalse
End Sub

Sub AlignCenterH()
    ActiveWindow.Selection.ShapeRange.Align msoAlignMiddles, msoFalse
End Sub

Sub DistributeHorizontally()
    ActiveWindow.Selection.ShapeRange.Distribute msoDistributeHorizontally, msoFalse
End Sub

Sub DistributeVertically()
    ActiveWindow.Selection.ShapeRange.Distribute msoDistributeVertically, msoFalse
End Sub

Sub SetShapeSize()
    Const Height As Long = 20
    Const Width As Long = 100
    If ActiveWindow.Selection.Type = ppSelectionShapes Then
        Dim sr As shape
        For Each sr In ActiveWindow.Selection.ShapeRange
            sr.Width = Width
            sr.Height = Height
        Next
    End If
End Sub

Sub SetFont()
    If ActiveWindow.Selection.Type = ppSelectionShapes Then
        Dim tr As TextRange
        For Each tr In ActiveWindow.Selection.TextRange
            With tr.Font
                .Name = "Segoe UI"
                .Size = 10
            End With
        Next
    End If
End Sub