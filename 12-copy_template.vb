' Test 12
' Copie un fichier template en entier, ajoute un style et sauvegarde

Public GeneratorPath As String
Public BasePath As String
Public BuildPath As String
Public DestPath As String
Public ProcessedDoc As Document

Private Function init()
    GeneratorPath = Options.DefaultFilePath(Path:=wdUserTemplatesPath) + "\generator"
    BasePath = GeneratorPath + "\src\base.dot"
    BuildPath = GeneratorPath + "\build"
    DestPath = BuildPath + "\copied.dot"
End Function

Private Function copyTemplate()
    FileCopy BasePath, DestPath
End Function

Private Function openCopiedDoc()
    Set ProcessedDoc = Documents.Open(FileName:=DestPath, Visible:=True)
End Function

Private Function addStyleToCopied()
    Set myStyle = ProcessedDoc.Styles.Add(Name:="StyleAjoute", _
    Type:=wdStyleTypeCharacter)
    With myStyle.Font
        .Bold = True
        .Italic = True
        .Name = "Arial"
        .Size = 12
    End With
End Function

Sub test()
    Call init
    Call copyTemplate
    Call openCopiedDoc
    Call addStyleToCopied
    ProcessedDoc.Save
    ProcessedDoc.Close
End Sub
