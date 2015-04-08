' Test 3
' Cette macro génère automatiquement un menu "Start Lodel" traduit dans la langue de Word

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

Private Sub startLodelMenu()
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
		.OnAction = "StartRevuesOrgDefault"
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
		.OnAction = "StartRevuesOrgEn"
	End With

	Set subMenuItem = subSubMenu.Controls.Add(Type:=msoControlButton)
	With subMenuItem
		.Caption = "Français" + showIfDefaultLang(" (par défaut)", "fr")
		.OnAction = "StartRevuesOrgFr"
	End With
	
	Set subMenuItem = subMenu.Controls.Add(Type:=msoControlButton)
	With subMenuItem
		.BeginGroup = True
		.Caption = trad("showComplet")
		.OnAction = "showComplet"
	End With
End Sub

'Sub AutoExec()
Sub test()
	wordLang = getWordLang()
    Call startLodelMenu
End Sub