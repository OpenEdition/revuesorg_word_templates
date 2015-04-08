' Test 1
' Macro simple qui attache le modèle et les macros au document en cours

Sub AutoExec()
    Application.CommandBars("LodelStart").Visible = True
End Sub

Sub LodelStart()
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
    AddIns( _
        Options.DefaultFilePath(Path:=wdUserTemplatesPath) + "\macros_revuesorg_win.dot" _
        ).Installed = True
    ActiveDocument.UpdateStylesOnOpen = True
    ActiveDocument.AttachedTemplate = Options.DefaultFilePath(Path:=wdUserTemplatesPath) + "\revuesorg_fr.dot"
    ActiveDocument.XMLSchemaReferences.AutomaticValidation = True
    ActiveDocument.XMLSchemaReferences.AllowSaveAsXMLWithoutValidation = False
End Sub