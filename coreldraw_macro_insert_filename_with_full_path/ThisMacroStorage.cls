VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisMacroStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub GlobalMacroStorage_DocumentBeforeSave(ByVal Doc As Document, ByVal SaveAs As Boolean, ByVal FileName As String)

Const LabelName As String = "DocLabel" ' The special name to mark the text label with
    Dim p As Page
    Dim CurPage As Page
    Dim s As Shape
    Dim CurSelection As ShapeRange
    
    ' Remember the current page
    Set CurPage = Doc.ActivePage
    ' Remember the current selection
    Set CurSelection = Doc.SelectionRange
    ' Go through every page and check for the labels
    For Each p In Doc.Pages
        ' Find the text label by object name
        Set s = p.FindShape(LabelName, cdrTextShape)
        If s Is Nothing Then
            ' The label doesn't exist. Create it
            Set s = p.ActiveLayer.CreateArtisticText(p.SizeWidth / 2, 0.2, FileName, Font:="Arial", Size:=10, Alignment:=cdrCenterAlignment)
            ' Set the special name to it
            s.Name = LabelName
        ElseIf SaveAs Then
            ' The label exists, but we are saving to another file. Update the label
            s.Text.Story = FileName
        End If
    Next p
    ' Restore the current page
    CurPage.Activate
    ' Restore the selection
    CurSelection.CreateSelection

End Sub

