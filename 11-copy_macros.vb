' Test 11
' Copier une macro d'un doc à un autre
' FIXME: Ne fonctionne pas à tout les coups !?? Ca semble venir de la gestion des autorisations de Word.

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

Private Function copyMacro(macroName As String)
    ' Pour une raison étrange Word envoie une erreur ici alors que la macro est correctement copiée.
    ' J'utilise 'On Error Resume Next' pour éviter l'erreur.
    On Error Resume Next
    Application.OrganizerCopy Source:= _
        BasePath, _
        Destination:=ProcessedDoc, Name:=macroName, Object:= _
        wdOrganizerObjectProjectItems
End Function

Sub test()
    Call init
    Call createNewDoc
    copyMacro("macroACopier")
    ProcessedDoc.Save
End Sub
