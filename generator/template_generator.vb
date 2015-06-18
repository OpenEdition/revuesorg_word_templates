'╔═══════════════════════════════════════════════════════╗
'║                 template_generator.vb                 ║
'║                 =====================                 ║
'║ Génération et traduction automatisées de modèles Word ║
'║ https://github.com/brrd/revuesorg_word_templates      ║
'╚═══════════════════════════════════════════════════════╝

' Déclarations
' ==========

Const TOOLBARNAME As String = "LodelStyles"
Const MACRONAME As String = "ApplyLodelStyle"

' Utilisé pour gérer les fichiers INI
Declare Function GetPrivateProfileString Lib "kernel32" Alias _
    "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
        ByVal lpKeyName As Any, ByVal lpDefault As String, _
        ByVal lpReturnedString As String, ByVal nSize As Long, _
        ByVal lpFileName As String) As Long

Public GeneratorPath As String
Public IniPath As String
Public AnsiIniPath As String
Public EnumerationsPath As String
Public AnsiEnumerationsPath As String
Public BasePath As String
Public BuildPath As String
Public TmpPath As String
Public DestLanguages() As String
Public LogFilePath As String
Public ProcessedDoc As Document

Private Function init()
    GeneratorPath = Options.DefaultFilePath(Path:=wdUserTemplatesPath) + "\generator"
    IniPath = GeneratorPath + "\src\translations.ini"
	AnsiIniPath = GeneratorPath + "\tmp\translations-ansi.ini"
    EnumerationsPath = GeneratorPath + "\utils\enumerations.ini"
    AnsiEnumerationsPath = GeneratorPath + "\tmp\enumerations.ini"
    BasePath = GeneratorPath + "\src\base.dot"
    TmpPath = GeneratorPath + "\tmp"
    BuildPath = GeneratorPath + "\build"
    LogFilePath = BuildPath + "\log.txt"
    Call getLanguagesFromIni
End Function

' Fichiers INI
' ==========

' Convertir un fichier encodé en Unicode en ANSI
Private Function unicode2ansi(source As String, dest As String)
    Dim strText
    With CreateObject("ADODB.Stream")
        .Type = 2
        .Charset = "utf-8"
        .Open
        .LoadFromFile source
        strText = .ReadText(-1)
        .Position = 0
        .SetEOS
        .Charset = "_autodetect_all"
        .WriteText strText, 0
        .SaveToFile dest, 2
        .Close
    End With
End Function

' Retourne un tableau de codes de langues tiré du INI
' Usage :
    'Dim intCount As Integer
    'For intCount = LBound(DestLanguages) To UBound(DestLanguages)
    '    Debug.Print Trim(DestLanguages(intCount))
    'Next
Private Function getLanguagesFromIni()
    Dim strLanguages As String
    Dim intCount As Integer

    strLanguages = getIniValue("_configuration", "translateTo")
    DestLanguages() = Split(strLanguages, ", ")
End Function

' Lire les fichiers INI
Private Function GetSectionEntry(ByVal strSectionName As String, ByVal strEntry As String, ByVal strIniPath As String) As String
    Dim X As Long
    Dim sSection As String, sEntry As String, sDefault As String
    Dim sRetBuf As String, iLenBuf As Integer, sFileName As String
    Dim sValue As String
    On Error GoTo ErrGetSectionentry
    sSection = strSectionName
    sEntry = strEntry
    sDefault = ""
    sRetBuf = Strings.String$(256, 0) '256 null characters
    iLenBuf = Len(sRetBuf$)
    sFileName = strIniPath
    X = GetPrivateProfileString(sSection, sEntry, _
        "", sRetBuf, iLenBuf, sFileName)
    sValue = Strings.Trim(Strings.Left$(sRetBuf, X))
    If sValue <> "" Then
        GetSectionEntry = sValue
    Else
        GetSectionEntry = vbNullChar
    End If
ErrGetSectionentry:
    If Err <> 0 Then
        Err.Clear
        Resume Next
    End If
End Function

' Lire translations.ini
Private Function getIniValue(section As String, key As String)
    getIniValue = GetSectionEntry(section, key, AnsiIniPath)
End Function

' Lire enumerations.ini
Private Function getEnumeration(section As String, key As String)
    getEnumeration = GetSectionEntry(section, key, AnsiEnumerationsPath)
End Function

' Récupérer la valeur d'une clé linguistique (ex: fr.menu), sinon la valeur par défaut (ex: menu)
Private Function getTranslation(section As String, key As String, lang As String)
    Dim value As String
    value = getIniValue(section, lang + "." + key)
    If value = vbNullChar Then
        value = getIniValue(section, key)
    End If
    getTranslation = value
End Function

' File operations
' ===========

Private Function FileExists(ByVal FileToTest As String) As Boolean
   FileExists = (Dir(FileToTest) <> "")
End Function

Private Function DeleteFile(ByVal FileToDelete As String)
   If FileExists(FileToDelete) Then 'See above
      SetAttr FileToDelete, vbNormal
      Kill FileToDelete
   End If
End Function

Private Function copyAndOpen(src As String, dest As String)
    FileCopy src, dest
    Set ProcessedDoc = Documents.Open(FileName:=dest, Visible:=True) 'FIXME: maintenant on a une erreur quand on set Visible=False !??
End Function

Private Function saveAndClose(doc As Document)
    doc.Save
    doc.Close
End Function

Private Function closeAll()
    With Application
        .ScreenUpdating = False
        Do Until .Documents.Count = 0
            .Documents(1).Close SaveChanges:=wdPromptToSaveChanges
        Loop
    End With
End Function

' Log
' ==========

Private Function openLog()
    DeleteFile LogFilePath
    Open LogFilePath For Append As #1
	Print #1, "GENERATOR.DOT : LOG DES ERREURS RENCONTREES (" + Format(Now, "ddddd ttttt") + ")"
	Print #1, "================================="
End Function

Private Function writeLog(msg As String)
    Print #1, msg
End Function

Private Function closeLog()
    Close #1
End Function

' Traduction des styles
' ==========

Private Function translateStyles(lang As String)
    Dim baseDocument As Document
    Dim id As String
    Dim newName As String
    Dim wordId As String

    writeLog ""
	writeLog "# Traduction des styles"

    Set baseDocument = Documents.Open(FileName:=BasePath, Visible:=False) ' TODO : a n'ouvrir qu'un fois pour toutes
    For Each Style In baseDocument.Styles
        If Style.BuiltIn = False Then
            id = Style.NameLocal
            newName = getTranslation(id, "style", lang)
            If newName <> vbNullChar Then
                If Not styleExists(newName) Then ProcessedDoc.Styles(id).NameLocal = newName
			Else
                If defaultName = vbNullChar Then
                    writeLog "Le style '" + id + "' n'a pas pu être traduit en langue " + lang
                End If
            End If
        Else
            wordId = getIniValue(id, "wordId")
            ' Traduire les builtInStyles pour ce modèle. N'est indispensable car Word les traduit automatiquement, c'est pourquoi on n'enregistre pas d'erreur si pas de traduction
            If wordId <> vbNullChar Then
                styleName = getTranslation(id, "style", lang)
                if styleName <> vbNullChar Then
                    ProcessedDoc.Styles(wordId).NameLocal = styleName
                End If
            End If
        End If
    Next Style

    ' Traduire les builtInStyles pour ce modèle. N'est indispensable car Word les traduit automatiquement, c'est pourquoi on n'enregistre pas d'erreur si pas de traduction
    If wordId <> vbNullChar Then
        styleName = getTranslation(id, "style", lang)
        if styleName <> vbNullChar Then
            ProcessedDoc.Styles(wordId).NameLocal = styleName
        End If
    End If

    baseDocument.Close
    ' Supprimer tous les styles préfixés ($) résiduels
    Call cleanPrefixedStyles
	writeLog "> Terminé"
End Function

' Nettoyer les styles prefixés avec $ à la fin du traitement (il s'agit des styles qui n'ont pas de traduction par défaut)
Private Function cleanPrefixedStyles()
    Dim firstChar As String
    For Each sty In ProcessedDoc.Styles
        If sty.BuiltIn = False Then
            firstChar = Left(sty.NameLocal, 1)
            If firstChar = "$" Then sty.Delete
        End If
    Next sty
End Function

Private Function styleExists(styleName As String) As Boolean ' TODO: doc en param
    Dim MyStyle As Word.Style
    On Error Resume Next
    Set MyStyle = ProcessedDoc.Styles(styleName)
    styleExists = Not MyStyle Is Nothing
End Function

' Keybindings (fonctions)
' ==========

' Génère un keyCode reconnu par Word à partir d'une chaîne du type "Ctrl+Alt+A"
' Voir le test 18 pour plus de détails
Private Function getKeyCode(keyString As String)
    Dim keys() As String
    Dim i As Integer
    Dim key As String
    Dim keyCode As String
    Dim sum As Long
    sum = 0
    keys = Split(keyString, "+")
    For i = 0 To UBound(keys)
        key = Trim(keys(i))
        keyCode = getEnumeration("keys", key)
        If (keyCode <> vbNullChar) Then
            sum = sum + CInt(keyCode)
        Else
            writeLog "Erreur : '" + key + "' n'est pas une touche valide"
            getKeyCode = 0
        End If
    Next i
    getKeyCode = sum
End Function

Private Function addStyleKeyBinding(styleName As String, keyString As String)
    Dim keyCode As Long
    keyCode = getKeyCode(keyString)
    If keyCode = 0 Then
        Exit Function
    End If
    ' Dans le cas d'un identifiant numerique il faut retrouver le nom du style
    If IsNumeric(styleName) Then
        styleName = ProcessedDoc.Styles(CInt(styleName)).NameLocal
    End If
    CustomizationContext = ProcessedDoc
    KeyBindings.Add KeyCategory:=wdKeyCategoryStyle, _
        Command:=styleName, _
        keyCode:=keyCode
End Function

' Une fonction pour supprimer tous les raccourcis clavier utilisateur d'un template
Private Function clearAllKeybindings()
    CustomizationContext = ProcessedDoc
    KeyBindings.ClearAll
End Function

' Traduction des menus et assignation des keybindings
' ==========

' Barre d'outils
' Fonction principale de traitement de la barre d'outil
Private Function processToolbar(lang As String)
	Dim Cmdbar As CommandBar
	Dim Ctl As CommandBarControl

	writeLog ""
	writeLog "# Traduction du menu et assignation des raccourcis clavier"

	Application.ScreenUpdating = False
	For Each Cmdbar In Application.CommandBars ' TODO: peut-être mieux d'utiliser la méthode .findControl() ?
	If Cmdbar.Name = TOOLBARNAME Then
		For Each Ctl In Cmdbar.Controls
			processSubmenu Ctl, lang
		Next Ctl
	End If
	Next Cmdbar
	writeLog "> Terminé"
End Function

' Fonction récursive de traitement des sous menus et assignation des raccourcis clavier
Private Function processSubmenu(submenu, lang As String)
    Dim menuName As String
    Dim styleName As String
    Dim menuId As String
	Dim wordId As String
    Dim key As String

    ' Traduire le caption du menu
    menuName = getTranslation(submenu.caption, "menu", lang)
    If menuName <> vbNullChar Then
        submenu.caption = menuName
    Else
        writeLog "Le sous-menu '" + submenu.caption + "' n'a pas pu être traduit en langue " + lang
    End If

    ' On boucle sur les enfants du menu'
    For Each Ctl In submenu.Controls
        menuId = Ctl.caption

        ' Si Ctl est un sous menu alors appel recursif de la fonction
        If Ctl.Type = msoControlPopup Then
            processSubmenu Ctl, lang
        ' Si Ctl est un control...
        ElseIf Ctl.Type = msoControlButton Then

            ' Traduire le caption du controle
            menuName = getTranslation(menuId, "menu", lang)
            wordId = getIniValue(menuId, "wordId")
            If menuName <> vbNullChar Then
                Ctl.caption = menuName
            ElseIf wordId <> vbNullChar Then ' Si la traduction d'un menu associé à un style natif n'est pas donnée alors on cherche la traduction de Word
                Ctl.caption = ProcessedDoc.Styles(wordId).NameLocal
            Else
                writeLog "Le bouton '" + menuId + "' n'a pas pu être traduit en langue " + lang
            End If

            ' La suite ne concerne pas les boutons qui contiennent un lien hypertexte
            If Ctl.HyperlinkType <> msoCommandBarButtonHyperlinkOpen Then

                ' Attacher la macro d'application de styles
                Ctl.OnAction = MACRONAME

                ' Assigner le parameter qui sera transmis a la macro d'application de styles
				If wordId <> vbNullChar Then
					Ctl.parameter = wordId
				Else
                    styleName = getTranslation(menuId, "style", lang)
					If styleName = vbNullChar Then
						styleName = Ctl.caption
					End If
					If Not styleExists(styleName) Then
						writeLog "Le style '" + styleName + "' associé au bouton '" + menuId + "' en langue " + lang + " n'existe pas. L'utilisation de ce bouton produira une erreur"
					End If
					Ctl.parameter = styleName
				End If

                ' Assigner le keybinding s'il existe
                ' Remarque : on n'ecrit rien dans le log concernant les keybindings pour éviter de le poluer
                key = getTranslation(menuId, "key", lang)
                If key <> vbNullChar Then
                    ' Assigner
                    If wordId <> vbNullChar Then
                        addStyleKeyBinding wordId, key
                    Else
                        addStyleKeyBinding styleName, key
                    End If
                    ' Ajouter la mention du raccourci dans le menu
                    Ctl.caption = Ctl.caption + vbTab + " <" + key + ">"
                End If
            End If
        End If
    Next Ctl
End Function

' Subs exposées
' ==========

' Ajouter un bouton à Word pour lancer runGenerator()
Sub AutoExec()
    Dim menuBar As CommandBar
    Dim menuItem As CommandBarControl
    Set menuBar = CommandBars.Add(menuBar:=False, Position:=msoBarTop, Name:="Template generator", Temporary:=True)
    menuBar.Visible = True
    Set menuItem = menuBar.Controls.Add(Type:=msoControlButton)
    With menuItem
        .Caption = "Générer les modèles traduits"
        .OnAction = "runGenerator"
        .Style = msoButtonCaption
    End With
End Sub

' Lancer la génération des modèles
Sub runGenerator()
    Dim currentLang As String
    Dim user As Integer
    ' Demande de confirmation pour la fermeture des documents ouverts dans Word
    If Application.Documents.Count <> 0 Then
        user = MsgBox("L'exécution du générateur de modèles va entrainer la fermeture tous les documents actuellement ouverts dans Word. L'enregistrement sera proposé pour tous les documents qui n'ont pas été sauvegardés.", vbOKCancel + vbQuestion, "Génération des modèles")
        If user <> vbOK Then
            Exit Sub
        End If
    End If
    Call closeAll
    ' Initialisation
    Call init
    Call openLog
    ' Convertir l'encodage des fichiers INI
	unicode2ansi IniPath, AnsiIniPath
    unicode2ansi EnumerationsPath, AnsiEnumerationsPath
    ' Copie et préparation de base.dot
    copyAndOpen BasePath, TmpPath + "\base.dot"
    Call clearAllKeybindings
    saveAndClose ProcessedDoc
    ' Création d'un modele par langue déclarée : traduction des styles, génération de la toolbar, attribution des actions et des keybindings
    For intCount = LBound(DestLanguages) To UBound(DestLanguages)
        currentLang = Trim(DestLanguages(intCount))
        If currentLang <> "" Then
            writeLog ""
            writeLog "╔═══════════════════════════════════════╗"
            writeLog "║ Generation du modele revuesorg_" + currentLang + ".dot ║"
            writeLog "╚═══════════════════════════════════════╝"
            copyAndOpen TmpPath + "\base.dot", BuildPath + "\revuesorg_" + currentLang + ".dot"
            translateStyles currentLang
            processToolbar currentLang
            saveAndClose ProcessedDoc
        End If
    Next
    ' Fin
    Call closeLog
    MsgBox "Les modèles ont été générés dans le dossier generator/build/." + Chr(10) + Chr(10) + "Merci de vérfier qu'aucune erreur n'a été rencontrée en consultant le journal generator/build/log.txt.", 64, "Opération terminée avec succès"
End Sub
