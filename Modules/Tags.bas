Attribute VB_Name = "Tags"
Option Explicit

Private TagFile As String
Private Const VimCommand As String = "gvim "

Public Sub ExportTags()
    Dim sl As Slide
    Dim shp As Shape
    Dim SlideIndex As Integer
    Dim ShapeIndex As Integer
    
    Open GetTagFile() For Output As #1
        SlideIndex = 1
        
        For Each sl In ActiveWindow.Presentation.Slides
            SlideIndex = SlideIndex + 1
            ShapeIndex = 1
            
            For Each shp In sl.Shapes
                ShapeIndex = ShapeIndex + 1
                
                Write #1, SlideIndex, ShapeIndex, shp.TextFrame2.TextRange.Text
            Next shp
        Next sl
    Close #1
    
    Shell GetTagFile(), vbNormalFocus
End Sub

Public Sub ImportTags()
    Dim SlideIndex As Integer
    Dim ShapeIndex As Integer
    Dim Text As String
    
    Open GetTagFile() For Input As #1
        Do While Not EOF(1)
            Input #1, SlideIndex, ShapeIndex, Text
            ActiveWindow.Presentation.Slides(SlideIndex - 1).Shapes(ShapeIndex - 1).TextFrame2.TextRange.Text = Text
        Loop
    Close #1
End Sub

Private Function GetTagFile() As String
    If TagFile = "" Then
        TagFile = Environ("TEMP") & "\tags.zen"
    End If
    
    GetTagFile = TagFile
End Function
