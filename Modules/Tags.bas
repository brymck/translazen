Attribute VB_Name = "Tags"
Option Explicit

Private TagFile As String
Private Const VimCommand As String = "gvim "
Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const TextData As Integer = 2
Private Const CreateOverwrite As Integer = 2
Private Const ReadLine As Integer = -2
Private Const TagSeparator As String = "--------"

Public Sub ExportTags()
    Dim sl As Slide
    Dim shp As Shape
    Dim BinaryStream As Object
    
    Set BinaryStream = CreateObject("ADODB.Stream")
    With BinaryStream
        .Type = TextData
        .Charset = "utf-8"
        .Open
        
        For Each sl In ActiveWindow.Presentation.Slides
            For Each shp In sl.Shapes
                .WriteText shp.TextFrame2.TextRange.Text
                .WriteText vbCrLf
                .WriteText TagSeparator
                .WriteText vbCrLf
            Next shp
        Next sl
        .SaveToFile GetTagFile(), CreateOverwrite
    End With
    
    ShellExecute 0, vbNullString, """" & GetTagFile() & """", vbNullString, vbNullString, vbNormalFocus
End Sub

Public Sub ImportTags()
    Dim sl As Slide
    Dim shp As Shape
    Dim BinaryStream As Object
    Dim PreviousLines As String
    Dim CurrentLine As String
    Dim FirstLine As Boolean
    
    Set BinaryStream = CreateObject("ADODB.Stream")
    With BinaryStream
        .Type = TextData
        .Charset = "utf-8"
        .Open
        .LoadFromFile GetTagFile()
        For Each sl In ActiveWindow.Presentation.Slides
            For Each shp In sl.Shapes
                ' Write current line unless it's a tag separator
                PreviousLines = vbNullString
                CurrentLine = .ReadText(ReadLine)
                FirstLine = True
                
                While CurrentLine <> TagSeparator
                    ' Add a line break unless it's the first string
                    If FirstLine Then
                        FirstLine = False
                    Else
                        PreviousLines = PreviousLines & vbCrLf
                    End If
                    
                    PreviousLines = PreviousLines & CurrentLine
                    CurrentLine = .ReadText(ReadLine)
                Wend
                
                shp.TextFrame2.TextRange.Text = PreviousLines
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
