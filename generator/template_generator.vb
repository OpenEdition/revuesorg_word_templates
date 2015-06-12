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

    strLanguages = getIniValue("_configuration", "translateTo") ' TODO: Harmonsier l'écriture des variables
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
    Set ProcessedDoc = Documents.Open(FileName:=dest, Visible:=True) ' TODO: Visible := False
End Function

Private Function saveAndClose(doc As Document)
    doc.Save
    doc.Close
End Function

Private Function closeAll() ' FIXME: risque de poser probleme si on utilise le modele en tant que modele. Plutôt associé en macro ?
    With Application
        .ScreenUpdating = False
        Do Until .Documents.Count = 0
            .Documents(1).Close SaveChanges:=wdDoNotSaveChanges
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
	Print #1, ""
End Function

Private Function writeLog(msg As String)
    Print #1, msg
End Function

Private Function closeLog()
    Close #1
End Function

' Traduction des styles
' ==========
' 1. Renommer les styles déjà présents dans base.dot d'après le fichier INI = BaseStyle. C'est l'entrée "style" sans prefixe de chaque section qui est utilisée (=français).
' 2. Pour chaque langue, regarder si une traduction existe dans le INI. Si elle existe et aucun style à son nom n'existe alors on le créée. Le style créé est hérité du BaseStyle.
' TODO: 1 et 2 à modifier, on n'utilise plus le défaut

Private Function renameBaseStylesAndTranslate()
	Dim styleId As String
	Dim newName As String
    Dim intCount As Integer
    Dim currentLang As String
	Dim baseDocument As Document
	Dim hasDefaultTranslation As Boolean

	writeLog "╔═══════════════════════╗"
	writeLog "║ TRADUCTION DES STYLES ║"
	writeLog "╚═══════════════════════╝"
	writeLog ""

	' Afin de ne pas créer d'interférences, on boucle sur les styles de base.dot et on modifie ceux d'ActiveDocument
	Set baseDocument = Documents.Open(FileName:=BasePath, Visible:=False)
    For Each Style In baseDocument.Styles
        If Style.BuiltIn = False Then
			styleId = Style.NameLocal
            ' Renommer le BaseStyle
            newName = getIniValue(styleId, "style")
            If newName <> vbNullChar Then
                ProcessedDoc.Styles(styleId).NameLocal = newName
				hasDefaultTranslation = true
			Else
				newName = styleId
				hasDefaultTranslation = false
            End If
            ' Le dupliquer en autant de traductions que nécessaire
            For intCount = LBound(DestLanguages) To UBound(DestLanguages)
                currentLang = Trim(DestLanguages(intCount))
                translateStyle newName, styleId, currentLang, hasDefaultTranslation
            Next
        End If
    Next Style
	baseDocument.Close
	writeLog "Traduction des styles terminée."
End Function

Private Function translateStyle(baseStyleName As String, styleId As String, lang As String, hasDefaultTranslation As Boolean) ' TODO: doc en param
    Dim translatedName As String
    Dim key As String
    Dim styleAdded As Style

    key = lang + ".style"
    translatedName = getIniValue(styleId, key)
    If translatedName = vbNullChar Or styleExists(translatedName) Then
		If hasDefaultTranslation = False Then
			writeLog "Le style '" + baseStyleName + "' n'a pas pu être traduit en langue " + lang
		End If
        Exit Function
    Else
        Set styleAdded = ProcessedDoc.Styles.Add(Name:=translatedName, _
            Type:=wdStyleTypeParagraph)
        styleAdded.baseStyle = baseStyleName
    End If
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
        styleName = ActiveDocument.Styles(CInt(styleName)).NameLocal
    End If
    CustomizationContext = ActiveDocument
    KeyBindings.Add KeyCategory:=wdKeyCategoryStyle, _
        Command:=styleName, _
        keyCode:=keyCode
End Function

' Une fonction pour supprimer tous les raccourcis clavier utilisateur d'un template
Private Function clearAllKeybindings()
    CustomizationContext = ActiveDocument
    KeyBindings.ClearAll
End Function

' Traduction des menus et assignation des keybindings
' ==========
' TODO: ajouter une traduction des builtInStyles dans chacun des modeles linguisitiques. Ca sert pas a grand chose car les builtIn prennent la langue de Word, mais c'est plus propre car harmonisé

' Barre d'outils
' Fonction principale de traitement de la barre d'outil
Private Function processToolbar(lang As String)
	Dim Cmdbar As CommandBar
	Dim Ctl As CommandBarControl

	writeLog ""
	writeLog "╔═════════════════════════╗"
	writeLog "║ TRADUCTION DU MENU [" + lang + "] ║"
	writeLog "╚═════════════════════════╝"
	writeLog ""

	Application.ScreenUpdating = False
	For Each Cmdbar In Application.CommandBars ' TODO: peut-être mieux d'utiliser la méthode .findControl() ?
	If Cmdbar.Name = TOOLBARNAME Then
		For Each Ctl In Cmdbar.Controls
			processSubmenu Ctl, lang
		Next Ctl
	End If
	Next Cmdbar
	writeLog "Traduction du menu en langue " + lang + " terminée."
End Function

' Fonction récursive de traitement des sous menus et assignation des raccourcis clavier
Private Function processSubmenu(submenu, lang As String)
    Dim menuName As String
    Dim styleName As String
    Dim menuId As String
	Dim wordId As String
    Dim key As String

    ' Traduire le caption du menu
    menuName = getIniValue(submenu.caption, lang + ".menu")
    If menuName = vbNullChar Then
        menuName = getIniValue(submenu.caption, "menu")
    End If
    If menuName <> vbNullChar Then
        submenu.caption = menuName
    Else
        writeLog "Le sous-menu '" + submenu.caption + "' n'a pas pu être traduit en langue " + lang
    End If

    ' On boucle sur les enfants du menu'
    For Each Ctl In submenu.Controls
        menuId = Ctl.caption

        ' Attacher les actions
        If Ctl.Type = msoControlButton Then
            If Ctl.HyperlinkType <> msoCommandBarButtonHyperlinkOpen Then ' Ne pas écraser les boutons de liens hypertextes
                Ctl.OnAction = MACRONAME
            End If
        ElseIf Ctl.Type = msoControlPopup Then ' FIXME : il faut grouper ça avec If Ctl.Type = msoControlButton Then ligne 356 sans quoi la fonction continue après processSubmenu
            ' Récursif sur les sous menus
            processSubmenu Ctl, lang
        End If

        ' Traduire le caption du controle
        menuName = getIniValue(menuId, lang + ".menu")
        If menuName = vbNullChar Then
            menuName = getIniValue(menuId, "menu")
        End If
        If menuName <> vbNullChar Then
            Ctl.caption = menuName
        Else
            writeLog "Le bouton '" + menuId + "' n'a pas pu être traduit en langue " + lang
        End If

        ' Assigner .parameter (qui doit être le nom du style à appliquer)
        If Ctl.Type = msoControlButton Then
			If Ctl.HyperlinkType <> msoCommandBarButtonHyperlinkOpen Then ' Ne pas écraser les boutons de liens hypertextes
				' Si possible on utilise les identifiants numériques de word pour les styles natifs. Voir : https://msdn.microsoft.com/en-us/library/bb237495%28v=office.12%29.aspx
				wordId = getIniValue(menuId, "wordId")
				If wordId <> vbNullChar Then
					Ctl.parameter = wordId
				Else
					styleName = getIniValue(menuId, lang + ".style")
					If styleName = vbNullChar Then
						styleName = getIniValue(menuId, "style")
					End If
					If styleName = vbNullChar Then
						styleName = Ctl.caption
					End If
					If Not styleExists(styleName) Then
						writeLog "Le style '" + styleName + "' associé au bouton '" + menuName + "' en langue " + lang + " n'existe pas. L'utilisation de ce bouton produira une erreur"
					End If
					Ctl.parameter = styleName
				End If

                ' Assigner le keybinding s'il existe
                ' Remarque : on n'ecrit rien dans le log concernant les keybindings pour éviter de le poluer
                ' TODO: serait bien de factoriser cette operation qui apparait plusieurs fois (fallback langue par defaut)
                key = getIniValue(menuId, lang + ".key")
                If key = vbNullChar Then
                    key = getIniValue(menuId, "key")
                End If
                If key <> vbNullChar Then
                    If wordId <> vbNullChar Then
                        addStyleKeyBinding wordId, key
                    Else
                        addStyleKeyBinding styleName, key
                    End If
                End If
			End If
        End If
    Next Ctl
End Function

' Tests

Sub TraduireLesStyles()
	Call closeAll
    Call init
    Call openLog
	' Convert INI encoding
	unicode2ansi IniPath, AnsiIniPath
    unicode2ansi EnumerationsPath, AnsiEnumerationsPath
	' TODO : idealement il faudrait juste copier les styles qui nous interessent de base.dot à styles.dot pour pas récupérer des indésirables. Le cas échéant, on pourrait copier comme doc de départ un doc vierge contenant la macro d'application des styles.
    copyAndOpen BasePath, TmpPath + "\styles.dot"
    Call renameBaseStylesAndTranslate
    Call clearAllKeybindings
    saveAndClose ProcessedDoc
End Sub

Sub en()
    Call TraduireLesStyles
    copyAndOpen TmpPath + "\styles.dot", BuildPath + "\revuesorg_en.dot"
    processToolbar "en"
End Sub

Sub Tout()
    Dim currentLang As String
    Call TraduireLesStyles
    For intCount = LBound(DestLanguages) To UBound(DestLanguages)
        currentLang = Trim(DestLanguages(intCount))
        If currentLang <> "" Then
            copyAndOpen TmpPath + "\styles.dot", BuildPath + "\revuesorg_" + currentLang + ".dot"
            processToolbar currentLang
            saveAndClose ProcessedDoc
        End If
    Next
    Call closeLog
End Sub
