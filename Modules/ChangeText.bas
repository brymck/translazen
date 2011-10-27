Attribute VB_Name = "ChangeText"
Option Explicit

Private Function HasText(ByRef shp As Shape) As Boolean
    If shp.TextFrame2.TextRange.Text <> vbNullString Then
        HasText = True
    Else
        HasText = False
    End If
End Function

Public Sub NextText(Optional ByVal UseNext As Boolean = True)
    Dim shp As Shape
    Dim CurrentSlide As Slide
    Dim Found As Boolean
    Dim CurrentShapeId As Long
    Dim LastTextShape As Shape
    
    Set CurrentSlide = ActiveWindow.View.Slide
    CurrentShapeId = ActiveWindow.Selection.ShapeRange.Id
    Found = False
    
    For Each shp In CurrentSlide.Shapes
        ' If we've already matched the shape, continue iterating through shapes until we find one
        ' with text
        If Found And UseNext And HasText(shp) Then
            shp.Select
            Exit Sub
        End If
        
        ' Set that we've found a shape
        If shp.Id = CurrentShapeId Then Found = True
        
        ' If using previous shape, we should have already found it
        If Not UseNext Then
            If Found Then
                ' Select last text shape if it's available
                On Error Resume Next
                LastTextShape.Select
                Exit Sub
            ElseIf HasText(shp) Then
                ' Set last text shape
                Set LastTextShape = shp
            End If
        End If
    Next shp
End Sub

Public Sub PreviousText()
    NextText False
End Sub
