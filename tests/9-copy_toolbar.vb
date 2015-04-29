' Test 9
' Copier une barre d'outils d'un doc à un autre (ici : de src/base.dot à un nouveau doc)

Public GeneratorPath As String
Public BasePath As String
Public BuildPath As String
Public ToolbarName As String
Public ProcessedDoc As Document

Private Function init()
    GeneratorPath = Options.DefaultFilePath(Path:=wdUserTemplatesPath) + "\generator"
    BasePath = GeneratorPath + "\src\base.dot"
    BuildPath = GeneratorPath + "\build"
    ToolbarName = "base_toolbar"
End Function

Private Function createNewDoc()
    Set ProcessedDoc = Documents.Add(BasePath, True, , True)
End Function

Private Function copyToolbar()
    CustomizationContext = ProcessedDoc
    Application.OrganizerCopy _
    Source:=BasePath, _
    Destination:=ProcessedDoc, Name:=ToolbarName, _
    Object:=wdOrganizerObjectCommandBars
    CommandBars(ToolbarName).Visible = True
End Function

Sub test()
    Call init
    Call createNewDoc
    Call copyToolbar
    ProcessedDoc.Save
End Sub
