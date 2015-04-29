' Test 10
' Copier tous les styles d'un document à un autre (ici : de src/base.dot à un nouveau doc)

Public GeneratorPath As String
Public BasePath As String
Public BuildPath As String
Public ProcessedDoc As Document

Private Function init()
    GeneratorPath = Options.DefaultFilePath(Path:=wdUserTemplatesPath) + "\generator"
    BasePath = GeneratorPath + "\src\base.dot"
    BuildPath = GeneratorPath + "\build"
End Function

Private Function createNewDoc()
    Set ProcessedDoc = Documents.Add(BasePath, True, , True)
End Function

Private Function copyStyles()
    ProcessedDoc.CopyStylesFromTemplate _
    Template:=BasePath
End Function

Sub test()
    Call init
    Call createNewDoc
    Call copyStyles
    ProcessedDoc.Save
End Sub
