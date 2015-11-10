' Test 5
' C'est une combinaison des Tests 2 et 3

Private wordLang As String

Private Function getWordLang() As String
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

Private Function trad(id As String, Optional lang As String = "")
	Dim key as String
	If lang = "" Then
		lang = wordLang
	End If
	key = lang + "." + id
    Select Case key
        Case "en.start"
            trad = "Start styling for Lodel"
		Case "fr.start"
			trad = "Démarrer le stylage pour Lodel"
        Case "en.options"
            trad = "Options"
		Case "fr.options"
            trad = "Options"
		Case "en.language"
			trad = "Language"
		Case "fr.language"
			trad = "Langue"
		Case "en.showComplet"
			trad = "Show additional styles"
		Case "fr.showComplet"
			trad = "Afficher les styles complémentaires"
        Case Else
            trad = "undefined"
    End Select
End Function

Function showIfDefaultLang (text As String, lang As String)
	If lang = wordLang Then
		showIfDefaultLang = text
	Else
		showIfDefaultLang = ""
	End If
End Function

Private Sub generateStartLodelMenu()
	Dim menuBar As CommandBar
	Dim subMenu As CommandBarControl
	Dim subSubMenu As CommandBarControl
	Dim menuItem As CommandBarControl
	Dim subMenuItem As CommandBarControl   

	Set menuBar = CommandBars.Add(menuBar:=False, Position:=msoBarTop, Name:="Start Lodel", Temporary:=True)
	menuBar.Visible = True

	Set menuItem = menuBar.Controls.Add(Type:=msoControlButton)
	With menuItem
		.Caption = trad("start")
		.OnAction = "startRevuesOrgDefault"
		.Style = msoButtonCaption
	End With

	Set subMenu = menuBar.Controls.Add(Type:=msoControlPopup)
	subMenu.BeginGroup = True
	subMenu.Caption = trad("options")
	
	Set subSubMenu = subMenu.Controls.Add(Type:=msoControlPopup)
	subSubMenu.Caption = trad("language")

	Set subMenuItem = subSubMenu.Controls.Add(Type:=msoControlButton)
	With subMenuItem
		.Caption = "English" + showIfDefaultLang(" (default)", "en")
		.OnAction = "startRevuesOrgEn"
	End With

	Set subMenuItem = subSubMenu.Controls.Add(Type:=msoControlButton)
	With subMenuItem
		.Caption = "Français" + showIfDefaultLang(" (par défaut)", "fr")
		.OnAction = "startRevuesOrgFr"
	End With
	
	Set subMenuItem = subMenu.Controls.Add(Type:=msoControlButton)
	With subMenuItem
		.BeginGroup = True
		.Caption = trad("showComplet")
		.OnAction = "showComplet"
	End With
End Sub

Sub doStart(Tpl As String)
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

Sub startRevuesOrgDefault()
    Dim Tpl As String
    Tpl = getTplPath(wordLang)
    doStart (Tpl)
End Sub

Sub startRevuesOrgFr()
    Dim Tpl As String
    Tpl = getTplPath("fr")
    doStart (Tpl)
End Sub

Sub startRevuesOrgEn()
    Dim Tpl As String
    Tpl = getTplPath("en")
    doStart (Tpl)
End Sub

Sub showComplet()
	MsgBox("Todo")
End Sub

Function getTplPath(Lang As String) As String
    getTplPath = Options.DefaultFilePath(Path:=wdUserTemplatesPath) + "/revuesorg_" + Lang + ".dot"
End Function

'Sub AutoExec()
Sub test()
	wordLang = getWordLang()
    Call generateStartLodelMenu
End Sub