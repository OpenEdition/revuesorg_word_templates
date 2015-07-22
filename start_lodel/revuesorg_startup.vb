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
		Case "en.startFullTemplate"
			trad = "Full template (advanced)"
		Case "fr.startFullTemplate"
			trad = "Modèle complet (avancé)"
        Case "en.templates"
			trad = "Other template"
		Case "fr.templates"
			trad = "Autre modèle"
        Case Else
            trad = "undefined"
    End Select
End Function

Private Function addTemplatesToMenu(subMenu As CommandBarControl)
    Dim tplPath As String
    Dim strFile As String
    Dim subMenuItem As CommandBarControl
    tplPath = Options.DefaultFilePath(Path:=wdUserTemplatesPath)
    strFile = Dir(tplPath + "\revuesorg_*.dot")
    Do While Len(strFile) > 0
        Set subMenuItem = subMenu.Controls.Add(Type:=msoControlButton)
        With subMenuItem
            .Caption = strFile
            .OnAction = "startOtherTemplate"
        End With
        strFile = Dir
    Loop
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
	subMenu.Caption = "+"

    Set subMenuItem = subMenu.Controls.Add(Type:=msoControlButton)
	With subMenuItem
		.BeginGroup = True
		.Caption = trad("startFullTemplate")
		.OnAction = "startFullTemplate"
	End With

	Set subSubMenu = subMenu.Controls.Add(Type:=msoControlPopup)
	subSubMenu.Caption = trad("templates")

    addTemplatesToMenu subSubMenu
End Sub

Sub doStart(tpl As String)
	Dim macro As String

    ' TODO: tester les paths sur OS X
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
    ActiveDocument.FormattingShowUserStyleName = True ' Afficher les noms de substitution quand on change le nom d'un style natif
    AddIns.Add FileName:=macro, Install:=True
    ActiveDocument.UpdateStylesOnOpen = True
    ActiveDocument.AttachedTemplate = tpl
    ActiveDocument.XMLSchemaReferences.AutomaticValidation = True
    ActiveDocument.XMLSchemaReferences.AllowSaveAsXMLWithoutValidation = False
    ActiveWindow.View.ShowBookmarks = True
End Sub

Sub startRevuesOrgDefault()
    Dim tpl As String
    tpl = Options.DefaultFilePath(Path:=wdUserTemplatesPath) + "\revuesorg_" + wordLang + ".dot"
    doStart tpl
End Sub

Sub startOtherTemplate()
    Dim ctlCBarControl  As CommandBarControl
    Dim tpl As String
    Set ctlCBarControl = CommandBars.ActionControl
    If ctlCBarControl Is Nothing Then Exit Sub
    tpl = Options.DefaultFilePath(Path:=wdUserTemplatesPath) + "\" + ctlCBarControl.caption
    doStart tpl
End Sub

Sub startFullTemplate()
    Dim tpl As String
    tpl = Options.DefaultFilePath(Path:=wdUserTemplatesPath) + "\revuesorg_complet_" + wordLang + ".dot"
    doStart tpl
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