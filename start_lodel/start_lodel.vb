' Start Lodel
' Generates a menu to handle automatically Revues.org templates and macros
' Installation : create a .dot file from this code and move it into the Word "Startup" folder

Private wordLang As String
Private os As String

Private Function getWordLang() As String
    ' https://msdn.microsoft.com/en-us/library/aa432635%28v=office.12%29.aspx
    Select Case Application.Language
    Case msoLanguageIDEnglishAUS, msoLanguageIDEnglishBelize, msoLanguageIDEnglishCanadian, msoLanguageIDEnglishCaribbean, msoLanguageIDEnglishIndonesia, msoLanguageIDEnglishIreland, msoLanguageIDEnglishJamaica, msoLanguageIDEnglishNewZealand, msoLanguageIDEnglishPhilippines, msoLanguageIDEnglishSouthAfrica, msoLanguageIDEnglishTrinidadTobago, msoLanguageIDEnglishUK, msoLanguageIDEnglishUS, msoLanguageIDEnglishZimbabwe
            getWordLang = "en"
        Case msoLanguageIDFrench, msoLanguageIDFrenchCameroon, msoLanguageIDFrenchCanadian, msoLanguageIDFrenchCotedIvoire, msoLanguageIDFrenchHaiti, msoLanguageIDFrenchLuxembourg, msoLanguageIDFrenchMali, msoLanguageIDFrenchMonaco, msoLanguageIDFrenchMorocco, msoLanguageIDFrenchReunion, msoLanguageIDFrenchSenegal, msoLanguageIDFrenchWestIndies, msoLanguageIDFranchCongoDRC, msoLanguageIDBelgianFrench, msoLanguageIDSwissFrench
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
		Case "en.showAdditionalTemplate"
			trad = "Show additional styles"
		Case "fr.showAdditionalTemplate"
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
		.Caption = trad("showAdditionalTemplate")
		.OnAction = "showAdditionalTemplate"
	End With
End Sub

Sub doStart(lang As String)
	Dim tpl As String
	Dim macro As String
	
    ' TODO: tester les paths sur OS X
	tpl = Options.DefaultFilePath(Path:=wdUserTemplatesPath) + "\revuesorg_" + lang + ".dot"
	macro = Options.DefaultFilePath(Path:=wdUserTemplatesPath) + "\macros_revuesorg_" + os + ".dot"
    
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
    AddIns.Add FileName:=macro, Install:=True
    ActiveDocument.UpdateStylesOnOpen = True
    ActiveDocument.AttachedTemplate = tpl
    ActiveDocument.XMLSchemaReferences.AutomaticValidation = True
    ActiveDocument.XMLSchemaReferences.AllowSaveAsXMLWithoutValidation = False
End Sub

Sub startRevuesOrgDefault()
    doStart (wordLang)
End Sub

Sub startRevuesOrgFr()
    doStart ("fr")
End Sub

Sub startRevuesOrgEn()
    doStart ("en")
End Sub

Sub showAdditionalTemplate()
	MsgBox("Todo")
End Sub

Sub AutoExec()
	' Testing OS: http://www.rondebruin.nl/mac/mac001.htm
	' TODO: A tester sur mac
    #If Mac Then
        os = "mac"
    #Else
        os = "win"
    #End If
	wordLang = getWordLang()
	Call generateStartLodelMenu    
End Sub