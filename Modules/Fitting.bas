Attribute VB_Name = "Fitting"
Option Explicit

Public Sub FitToShape()
    With ActiveWindow.Selection.ShapeRange.TextFrame2
        .AutoSize = msoAutoSizeTextToFitShape
    End With
End Sub
