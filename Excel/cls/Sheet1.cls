Sub Format()
    With Me
        With .PageSetup
            .PrintArea = Me.UsedRange.Address
            .Orientation = IIf(Me.UsedRange.Columns.Count < Me.UsedRange.Rows.Count, xlPortrait, xlLandscape)
            .PaperSize = xlPaperA4
            .Zoom = False 'Must be False to enable FitToPagesTall and FitToPagesWide
            .FitToPagesTall = False
            .FitToPagesWide = 1
            .HeaderMargin = 0
            .FooterMargin = 0
            .TopMargin = 0
            .BottomMargin = 0
            .LeftMargin = 0
            .RightMargin = 0
            .PrintGridlines = True
        End With

        With .Cells.Font
            .Name = "Segoe UI"
            .Size = 10
            .Rows.AutoFit
            .Columns.AutoFit
        End With
    End With
End Sub