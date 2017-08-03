Option Explicit

Private Sub Main()
    Dim d As Document
    Set d = ThisDocument

'    PrintPageSetup d
'    Dim s As Section
'    For Each s In d.Sections
'        PrintSectionPageSetup s
'    Next
'
'    PrintBodyHeaderFooterBookmarks d, False, False
'
'    PrintBodyText d
'    PrintHeaderText d.Sections(1), False, False
'    PrintFooterText d.Sections(1), False, False

    PrintBodyShapes d
    PrintHeaderFooterShapes d

    d.ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
End Sub

Private Sub RenameShapeToAmendment(ByVal s As Shape)
    If s.Name = "Amendment" Or s.Name = "Rectangle 5" Or s.Name = "Rectangle 12" Then
        s.Name = "Amendment"
        s.Fill.ForeColor = RGB(255, 255, 0)
    End If
End Sub

Private Sub PrintPageSetup(ByVal d As Document)
    Printf "DifferentFirstPageHeaderFooter = {0}", CBool(d.PageSetup.DifferentFirstPageHeaderFooter)
    Printf "OddAndEvenPagesHeaderFooter = {0}", CBool(d.PageSetup.OddAndEvenPagesHeaderFooter)
    ' DifferentFirstPageHeaderFooter
    ' https://msdn.microsoft.com/en-us/library/office/ff195626.aspx
    ' OddAndEvenPagesHeaderFooter
    ' https://msdn.microsoft.com/en-us/library/office/ff821984.aspx
End Sub

Private Sub PrintBodyHeaderFooterBookmarks(ByVal d As Document, ByVal overwrite As Boolean, ByVal insert As Boolean)
    Dim b As Bookmark
    For Each b In d.Bookmarks
        With b.Range
            If overwrite Then
                'Bookmark.Text is write-only. Reading the property returns an empty string.
                'Bookmark.Text is only a target to throw a string at.
                .Text = StringFormat("(Bookmark = {0})", b.Name)
            End If
            If insert Then
                .InsertBefore "(Before)"
                .InsertAfter "(After)"
            End If
            Printf "Bookmark = {0}", b.Name
        End With
    Next
End Sub

Private Sub PrintBodyText(ByVal d As Document)
    Debug.Print "Text =" & d.Range.Text
End Sub

Private Sub PrintBodyShapes(ByVal d As Document)
    Dim s As Shape
    For Each s In d.Shapes
        RenameShapeToAmendment s
        Printf "Body Shape = {0}", s.Name
    Next
End Sub

Private Sub PrintHeaderFooterShapes(ByVal d As Document)
    ' Setting wdSeekMainDocument to SeekView makes d.Application.Selection.HeaderFooter null.
    ' Setting one of the following 6 to SeekView causes an error when trying to assing to SeekView.
    ' wdSeekFirstPageHeader
    ' wdSeekEvenPagesHeader
    ' wdSeekFirstPageFooter
    ' wdSeekEvenPagesFooter
    ' wdSeekFootnotes
    ' wdSeekEndnotes

    ' Setting one of the following 4 to SeekView returns the same result.
    ' Namely, wdSeekPrimaryHeader returns shapes not only in the header but also shapes in the footer.
    ' Namely, wdSeekPrimaryFooter returns shapes not only in the footer but also shapes in the header.
    ' wdSeekPrimaryHeader
    ' wdSeekPrimaryFooter
    ' wdSeekCurrentPageHeader
    ' wdSeekCurrentPageFooter

    ' So setting wdSeekPrimaryHeader to SeekView is enough.
    d.ActiveWindow.ActivePane.View.SeekView = wdSeekPrimaryHeader

    Dim s As Shape
    If Not d.Application.Selection.HeaderFooter Is Nothing Then
        For Each s In d.Application.Selection.HeaderFooter.Shapes
            RenameShapeToAmendment s
            Printf "HeaderFooter Shape = {0}", s.Name
        Next
    End If
    d.ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
End Sub

Private Sub PrintSectionPageSetup(ByVal s As Section)
    Printf "(SectionIndex, DifferentFirstPageHeaderFooter) = ({0}, {1})", s.Index, CBool(s.PageSetup.DifferentFirstPageHeaderFooter)
    Printf "(SectionIndex, OddAndEvenPagesHeaderFooter) = ({0}, {1})", s.Index, CBool(s.PageSetup.OddAndEvenPagesHeaderFooter)
End Sub

Private Sub PrintHeaderText(ByVal s As Section, ByVal overwrite As Boolean, ByVal insert As Boolean)
    Dim h As HeaderFooter
    For Each h In s.Headers
        Dim e As String
        e = GetHeaderFooterEnum(h.Index)
        With h.Range
            If overwrite Then
                .Text = StringFormat("(Header = {0})", e)
            End If
            If insert Then
                .InsertBefore "(Before)"
                .InsertAfter "(After)"
            End If
            Printf "(SectionIndex, Header, Header.Text) = ({0}, {1}, {2})", s.Index, e, MyTrim(.Text)
        End With
    Next
End Sub

Private Sub PrintFooterText(ByVal s As Section, ByVal overwrite As Boolean, ByVal insert As Boolean)
    Dim f As HeaderFooter
    For Each f In s.Footers
        Dim e As String
        e = GetHeaderFooterEnum(f.Index)
        With f.Range
            If overwrite Then
                .Text = StringFormat("(Footer = {0})", e)
            End If
            If insert Then
                .InsertBefore "(Before)"
                .InsertAfter "(After)"
            End If
            Printf "(SectionIndex, Footer, Footer.Text) = ({0}, {1}, {2})", s.Index, e, MyTrim(.Text)
        End With
    Next
End Sub

Private Function GetHeaderFooterEnum(ByVal i As Long) As String
    Select Case i
        Case 1
            GetHeaderFooterEnum = "wdHeaderFooterPrimary"
        Case 2
            GetHeaderFooterEnum = "wdHeaderFooterFirstPage"
        Case 3
            GetHeaderFooterEnum = "wdHeaderFooterEvenPages"
        Case Else
            Err.Raise 93 'Invalid pattern string (http://support.microsoft.com/kb/146864)
    End Select
    ' WdHeaderFooterIndex Enumeration
    ' https://msdn.microsoft.com/en-us/library/office/ff839314.aspx
End Function

Private Function GetSeekViewEnum(ByVal i As Long) As String
    Select Case i
        Case 0
            GetSeekViewEnum = "wdSeekMainDocument"
        Case 1
            GetSeekViewEnum = "wdSeekPrimaryHeader"
        Case 2
            GetSeekViewEnum = "wdSeekFirstPageHeader"
        Case 3
            GetSeekViewEnum = "wdSeekEvenPagesHeader"
        Case 4
            GetSeekViewEnum = "wdSeekPrimaryFooter"
        Case 5
            GetSeekViewEnum = "wdSeekFirstPageFooter"
        Case 6
            GetSeekViewEnum = "wdSeekEvenPagesFooter"
        Case 7
            GetSeekViewEnum = "wdSeekFootnotes"
        Case 8
            GetSeekViewEnum = "wdSeekEndnotes"
        Case 9
            GetSeekViewEnum = "wdSeekCurrentPageHeader"
        Case 10
            GetSeekViewEnum = "wdSeekCurrentPageFooter"
        Case Else
            Err.Raise 93 'Invalid pattern string (http://support.microsoft.com/kb/146864)
    End Select
    ' WdSeekView enumeration
    ' https://msdn.microsoft.com/en-us/library/office/ff197738.aspx
End Function

Public Function MyTrim(ByVal s As String) As String
    MyTrim = Trim$(Trim$(Replace(Replace(Replace(s, vbCr, vbNullString), vbLf, vbNullString), vbCrLf, vbNullString)))
End Function

Public Function StringFormat(ByVal format As String, ParamArray args()) As String
    Dim i As Long
    For i = 0 To UBound(args)
        format = Replace(format, "{" & i & "}", args(i))
    Next
    StringFormat = format
End Function

Public Sub Printf(ByVal format As String, ParamArray args())
    Dim i As Long
    For i = 0 To UBound(args)
        format = Replace(format, "{" & i & "}", args(i))
    Next
    Debug.Print format
End Sub
