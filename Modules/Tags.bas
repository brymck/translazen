Attribute VB_Name = "Tags"
Option Explicit

Private TagFile As String
Private Const VimCommand As String = "gvim "
Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const ReadLine As Integer = -2

Public Sub ExportTags()
    Dim sl As Slide
    Dim shp As Shape
    Dim BinaryStream As Object
    
    Set BinaryStream = CreateObject("ADODB.Stream")
    With BinaryStream
        .Type = 2
        .Charset = "utf-8"
        .Open
        For Each sl In ActiveWindow.Presentation.Slides
            For Each shp In sl.Shapes
                .WriteText shp.TextFrame2.TextRange.Text
                .WriteText vbCrLf
            Next shp
        Next sl
        .SaveToFile GetTagFile(), 2
    End With
    
    ShellExecute 0, vbNullString, """" & GetTagFile() & """", vbNullString, vbNullString, vbNormalFocus
End Sub

Public Sub ImportTags()
    Dim sl As Slide
    Dim shp As Shape
    Dim BinaryStream As Object
    Dim ReadFile As String
    
    Set BinaryStream = CreateObject("ADODB.Stream")
    With BinaryStream
        .Type = 2
        .Charset = "utf-8"
        .Open
        .LoadFromFile GetTagFile()
        For Each sl In ActiveWindow.Presentation.Slides
            For Each shp In sl.Shapes
                shp.TextFrame2.TextRange.Text = .ReadText(ReadLine)
            Next shp
        Next sl
    End With
End Sub

Private Function GetTagFile() As String
    If TagFile = "" Then
        TagFile = Environ("TEMP") & "\tags.zen"
    End If
    
    GetTagFile = TagFile
End Function
