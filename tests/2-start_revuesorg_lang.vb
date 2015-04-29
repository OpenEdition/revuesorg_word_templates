' Test 2
' StartRevuesOrgDefault() détecte la langue de Word et attache le modèle et les macros correspondants au document en cours
' StartRevuesOrgFr() et StartRevuesOrgEn() attachent les versions linguistiques correspondantes

Sub DoStart(Tpl As String)
    If ActiveWindow.View.SplitSpecial = wdPaneNone Then
        ActiveWindow.ActivePane.View.Type = wdNormalView
    Else
        ActiveWindow.View.Type = wdNormalView
    End If
    ActiveWindow.StyleAreaWidth = CentimetersToPoints(3)
    ActiveDocument.ActiveWindow.View.ShowAll = True
    ActiveDocument.FormattingShowFont = True
    ActiveDocument.FormattingShowParagraph = True
    ActiveDocument.FormattingShowNumbering = True
    ActiveDocument.FormattingShowFilter = wdShowFilterStylesInUse
    AddIns.Add FileName:=Options.DefaultFilePath(Path:=wdUserTemplatesPath) + "\macros_revuesorg_win.dot", Install:=True
    ActiveDocument.UpdateStylesOnOpen = True
    ActiveDocument.AttachedTemplate = Tpl
    ActiveDocument.XMLSchemaReferences.AutomaticValidation = True
    ActiveDocument.XMLSchemaReferences.AllowSaveAsXMLWithoutValidation = False
End Sub

Sub StartRevuesOrgDefault()
    Dim Lang As String
    Dim Tpl As String
    Lang = getWordLang()
    Tpl = Options.DefaultFilePath(Path:=wdUserTemplatesPath) + "/revuesorg_" + Lang + ".dot"
    DoStart (Tpl)
End Sub

Sub StartRevuesOrgFr()
    Dim Tpl As String
    Tpl = GetTplPath("fr")
    DoStart (Tpl)
End Sub

Sub StartRevuesOrgEn()
    Dim Tpl As String
    Tpl = GetTplPath("en")
    DoStart (Tpl)
End Sub

Function getWordLang() As String
' https://technet.microsoft.com/en-us/library/cc287874%28v=office.12%29.aspx
    currentLanguageCode = Selection.LanguageID
    Select Case currentLanguageCode
        Case 1033
            getWordLang = "en"
        Case 1036
            getWordLang = "fr"
        Case Else
            getWordLang = "en" ' default
    End Select
End Function

Function GetTplPath(Lang As String) As String
    GetTplPath = Options.DefaultFilePath(Path:=wdUserTemplatesPath) + "/revuesorg_" + Lang + ".dot"
End Function