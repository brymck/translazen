Attribute VB_Name = "Buttons"
Option Explicit

Private Sub AddButtons()
    Dim cb As CommandBar
    
    On Error Resume Next
    Application.CommandBars("pp_yaku_zen").Delete
    On Error GoTo 0
    
    Set cb = Application.CommandBars.Add("pp_yaku_zen", msoBarTop, , True)

    With cb.Controls
        ' Add fit to shape
        With .Add(msoControlButton)
            .Caption = "&Fit to Shape"
            .OnAction = "FitToShape"
            .Style = msoButtonCaption
        End With
        
        ' Add regex search
        With .Add(msoControlButton)
            .Caption = "&Regex Search"
            .OnAction = "RegexSearch"
            .Style = msoButtonCaption
        End With
    End With
    cb.Visible = True
End Sub


Public Sub Auto_Open()
    AddButtons
End Sub
