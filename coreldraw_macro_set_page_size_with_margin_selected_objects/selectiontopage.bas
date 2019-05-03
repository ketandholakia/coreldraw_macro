Attribute VB_Name = "selectiontopage"
Option Explicit

Sub fitCanvas()

    Dim s As Shape
    Dim w As Double, h As Double
    
    Set s = ActiveSelection
    If s.Shapes.Count = 0 Then
        MsgBox "Please make a selection"
        Exit Sub
    End If
    s.GetSize w, h
    
    ActivePage.SizeHeight = h + 0.5
    ActivePage.SizeWidth = w + 0.5
    ActiveDocument.ReferencePoint = cdrBottomLeft
    s.SetPosition 0.25, 0.25

End Sub


