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
        If Found And UseNext Then
            If HasText(shp) Then
                shp.Select
                Exit Sub
            End If
        End If
        
        If shp.Id = CurrentShapeId Then Found = True
        
        If Not UseNext Then
            If Found Then
                LastTextShape.Select
            ElseIf HasText(shp) Then
                Set LastTextShape = shp
            End If
        End If
    Next shp
End Sub

Public Sub PreviousText()
    IncText False
End Sub
