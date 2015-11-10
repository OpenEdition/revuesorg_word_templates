'+-------------------------------------------------------+
'¦                 template_generator.vb                 ¦
'¦                 =====================                 ¦
'¦ Génération et traduction automatisées de modèles Word ¦
'¦ https://github.com/brrd/revuesorg_word_templates      ¦
'+-------------------------------------------------------+

' Déclarations
' ==========

' Chemins
Const ROOT As String = "\revuesorg_word_templates"
Const BASEDOT As String = "\src\base.dot"
Const TRANSLATIONS As String = "\src\translations.ini"
Const ENUMERATIONS As String = "\utils\enumerations.ini"
Const TMP = "\tmp"
Const TMPTRANSLATIONS As String = "\tmp\translations.tmp.ini"
Const TMPENUMERATIONS As String = "\tmp\enumerations.tmp.ini"
Const TMPBASEDOT As String = "\tmp\base.tmp.dot"
Const BUILD = "\build"
Const BUILDTEMPLATES = "\build\templates"
Const LOGTXT As String = "\build\generator_log.txt"

' Elements dans base.dot
Const TOOLBARNAME As String = "LodelStyles"
Const MACRONAME As String = "ApplyLodelStyle"

' Utilisé pour gérer les fichiers INI
Public Declare Function GetPrivateProfileString _
                            Lib "Kernel32" Alias "GetPrivateProfileStringW" _
                            (ByVal lpApplicationName As String, _
                             ByVal lpKeyName As String, _
                             ByVal lpDefault As String, _
                             lpReturnedString As Any, _
                             ByVal nSize As Long, _
                             ByVal lpFileName As String) As Long

' TODO: a transmettre en var
Public DestLanguages() As String
Public ProcessedDoc As Document

' Fichiers INI
' ==========

' Convertir un fichier encodé en UTF-8 en UTF-16
Private Function toUtf16(ByVal source As String, ByVal dest As String)
    Dim strText
    With CreateObject("ADODB.Stream")
        .Type = 2 'Specify stream type - we want To save text/string data.
        .Charset = "utf-8" 'Specify charset For the source text data.
        .Open 'Open the stream And write binary data To the object
        .LoadFromFile source
        strText = .ReadText(-1)
        .Position = 0
        .SetEOS
        .Charset = "utf-16"
        .WriteText strText, 0
        .SaveToFile dest, 2
        .Close
    End With
End Function

' Initialisation des langues dans la globale DestLanguages(). La fonction retourne un booléen selon le succès de la lecture du INI.
Private Function getLanguagesFromIni() As Boolean
    Dim strLanguages As String
    Dim intCount As Integer
    strLanguages = getIniValue("_configuration", "translateTo")
    If strLanguages <> vbNullChar Then
        DestLanguages() = Split(strLanguages, ", ")
        getLanguagesFromIni = True
    Else
        getLanguagesFromIni = False
    End If
End Function

' Lire les fichiers ini avec support d'Unicode. Le fichier source doit être encodé en UTF-16 LE.
' Voir : http://www.access-programmers.co.uk/forums/showthread.php?t=164136
Public Function profileGetItem(ByVal sSection As String, _
                                ByVal sKeyName As String, _
                                ByVal sInifile As String) As String
    Dim retval As Long
    Dim cSize As Long
    Dim Buf() As Byte
    ReDim Buf(254)
    cSize = 255
    Dim sDefValue As String
    sDefValue = ""
    retval = GetPrivateProfileString(StrConv(sSection, vbUnicode), _
                                        StrConv(sKeyName, vbUnicode), _
                                        StrConv(sDefValue, vbUnicode), _
                                        Buf(0), _
                                        cSize, _
                                        StrConv(sInifile, vbUnicode))
    If retval > 0 Then
        profileGetItem = Left(Buf, retval)
    Else
        profileGetItem = vbNullChar
    End If
End Function

' Lire translations.ini
Private Function getIniValue(ByVal section As String, ByVal key As String) As String
    getIniValue = profileGetItem(section, key, getAbsPath(TMPTRANSLATIONS))
End Function

' Lire enumerations.ini
Private Function getEnumeration(ByVal section As String, ByVal key As String) As String
    getEnumeration = profileGetItem(section, key, getAbsPath(TMPENUMERATIONS))
End Function

' Récupérer la valeur d'une clé linguistique (ex: fr.menu), sinon la valeur par défaut (ex: menu)
Private Function getTranslation(ByVal section As String, ByVal key As String, ByVal lang As String) As String
    Dim value As String
    value = getIniValue(section, lang + "." + key)
    If value = vbNullChar Then
        value = getIniValue(section, key)
    End If
    getTranslation = value
End Function

' File operations
' ===========

Private Function getAbsPath(Optional ByVal relPath As String = "") As String
    getAbsPath = Options.DefaultFilePath(Path:=wdUserTemplatesPath) + ROOT + relPath
End Function

Private Function fileExists(ByVal FileToTest As String) As Boolean
   fileExists = (Dir(FileToTest) <> "")
End Function

Private Function deleteFile(ByVal FileToDelete As String)
   If fileExists(FileToDelete) Then
      SetAttr FileToDelete, vbNormal
      Kill FileToDelete
   End If
End Function

Private Function folderExists(ByVal folderToTest As String) As Boolean
    folderExists = (Dir(folderToTest, vbDirectory) <> "")
End Function

Private Function deleteFolder(ByVal folderToDelete As String)
    If folderExists(folderToDelete) Then
        Kill folderToDelete + "\*.*"
        RmDir folderToDelete
    End If
End Function

Private Function createFolder(ByVal folderPath As String)
    If Not folderExists(folderPath) Then
        MkDir folderPath
    End If
End Function

Private Function copyAndOpen(ByVal src As String, ByVal dest As String)
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
    Dim logFile As String
    logFile = getAbsPath(LOGTXT)
    deleteFile logFile
    Open logFile For Append As #1
    Print #1, "GENERATOR : LOG DES ERREURS RENCONTREES (" + Format(Now, "ddddd ttttt") + ")"
    Print #1, "================================="
    Print #1, "Chemin sur le disque : " + getAbsPath(LOGTXT)
End Function

Private Function writeLog(ByVal msg As String)
    Print #1, msg
End Function

Private Function closeLog()
    Close #1
End Function

' Traduction des styles
' ==========

Private Function translateStyles(ByVal lang As String, ByVal isComplet As Boolean)
    Dim baseDocument As Document
    Dim id As String
    Dim newName As String

    writeLog ""
    writeLog "# Traduction des styles"

    Set baseDocument = Documents.Open(FileName:=getAbsPath(BASEDOT), Visible:=False) ' TODO : a n'ouvrir qu'un fois pour toutes
    For Each Style In baseDocument.Styles
        id = Style.NameLocal

        If Not (UCase(getIniValue(id, "complet")) = "TRUE" And Not isComplet) Then ' Ne traduire les styles complets que quand c'est demandé
            newName = getTranslation(id, "style", lang)
            If newName <> vbNullChar Then
                If Not styleExists(newName) Then ProcessedDoc.Styles(id).NameLocal = newName
            Else
                If defaultName = vbNullChar Then
                    writeLog "  Le style '" + id + "' n'a pas pu être traduit en langue " + lang + "."
                End If
            End If
        End If
    Next Style

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

Private Function styleExists(ByVal styleName As String) As Boolean ' TODO: doc en param
    Dim MyStyle As Word.Style
    On Error Resume Next
    Set MyStyle = ProcessedDoc.Styles(styleName)
    styleExists = Not MyStyle Is Nothing
End Function

' Keybindings (fonctions)
' ==========

' Génère un keyCode reconnu par Word à partir d'une chaîne du type "Ctrl+Alt+A"
' Voir le test 18 pour plus de détails
Private Function getKeyCode(ByVal keyString As String)
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
            writeLog "  Erreur : '" + key + "' n'est pas une touche valide" + "."
            getKeyCode = 0
        End If
    Next i
    getKeyCode = sum
End Function

Private Function addStyleKeyBinding(ByVal styleName As String, ByVal keyString As String)
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
Private Function processToolbar(ByVal lang As String, ByVal isComplet As String)
    Dim Cmdbar As CommandBar
    Dim Ctl As CommandBarControl

    writeLog ""
    writeLog "# Traduction du menu et assignation des raccourcis clavier"

    Application.ScreenUpdating = False
    For Each Cmdbar In Application.CommandBars ' TODO: peut-être mieux d'utiliser la méthode .findControl() ?
    If Cmdbar.Name = TOOLBARNAME Then
        For Each Ctl In Cmdbar.Controls
            processSubmenu Ctl, lang, isComplet
        Next Ctl
    End If
    Next Cmdbar
    writeLog "> Terminé"
End Function

' Fonction récursive de traitement des sous menus et assignation des raccourcis clavier
Private Function processSubmenu(submenu, ByVal lang As String, ByVal isComplet As Boolean)
    Dim menuName As String
    Dim styleName As String
    Dim menuId As String
    Dim key As String
    Dim hyperlink As String
    Dim builtIn As String

    ' Supprimer le menu s'il n'est pas censé être dans ce modèle (styles complets)
    If UCase(getIniValue(submenu.Caption, "complet")) = "TRUE" And Not isComplet Then
        submenu.Delete
        Exit Function
    End If

    ' Traduire le caption du menu
    menuName = getTranslation(submenu.Caption, "menu", lang)
    If menuName <> vbNullChar Then
        submenu.Caption = menuName
    Else
        writeLog "  Le sous-menu '" + submenu.Caption + "' n'a pas pu être traduit en langue " + lang + ". Le sous-menu a été supprimé du modèle."
        submenu.Delete
        Exit Function
    End If

    ' On boucle sur les enfants du menu'
    For Each Ctl In submenu.Controls
        menuId = Ctl.Caption

        ' Si Ctl est un sous menu alors appel recursif de la fonction
        If Ctl.Type = msoControlPopup Then
            processSubmenu Ctl, lang, isComplet
        ' Si Ctl est un control...
        ElseIf Ctl.Type = msoControlButton Then

            ' Traduire le caption du controle
            menuName = getTranslation(menuId, "menu", lang)
            If menuName <> vbNullChar Then
                Ctl.Caption = menuName
            Else
                Ctl.Delete
                writeLog "  Le bouton '" + menuId + "' n'a pas pu être traduit en langue " + lang + ". Le bouton a été supprimé du modèle."
                Goto NextCtl
            End If

            ' Assigner un lien hypertexte au bouton avec 'hyperlink'
            hyperlink = getTranslation(menuId, "hyperlink", lang)
            If hyperlink <> vbNullChar Then
                Ctl.HyperlinkType = msoCommandBarButtonHyperlinkOpen
                Ctl.TooltipText = hyperlink
            End If

            ' La suite ne concerne pas les boutons qui contiennent un lien hypertexte
            If Ctl.HyperlinkType <> msoCommandBarButtonHyperlinkOpen Then

                ' Assigner le parameter qui sera transmis a la macro d'application de styles
                styleName = getTranslation(menuId, "style", lang)
                If styleName = vbNullChar Then
                    styleName = Ctl.Caption
                End If
                ' On vérifie que le style correspondant existe bien avant de créer le bouton (sauf pour les styles natifs)
                builtIn = getTranslation(menuId, "builtIn", lang)
                If Not styleExists(styleName) And builtIn <> "true" Then
                    Ctl.Delete
                    writeLog "  Le style '" + styleName + "' associé au bouton '" + menuId + "' n'existe pas. Le bouton a été supprimé du modèle."
                    Goto NextCtl
                End If
                Ctl.Parameter = styleName

                ' Assigner le keybinding s'il existe
                ' Remarque : on n'ecrit rien dans le log concernant les keybindings pour éviter de le poluer
                key = getTranslation(menuId, "key", lang)
                If key <> vbNullChar Then
                    ' Assigner
                    addStyleKeyBinding styleName, key
                    ' Ajouter la mention du raccourci dans le menu
                    Ctl.Caption = Ctl.Caption + " <" + key + ">"
                End If
            End If
        End If
        NextCtl:
    Next Ctl
End Function

' Texte du modèle (informations de version)
' ==========

Private Function addTplInfo(tplName As String)
    Dim tplVersion As String
    Dim text As String
    tplVersion = getIniValue("_configuration", "version")
    If tplVersion <> vbNullChar Then
        text = tplName + " - version " + tplVersion + Chr(10) + "Info: http://maisondesrevues.org/"
    Else
        writeLog " Attention. La version des modèles traduits n'est pas spécifiée dans la fichier de configuration."
        text = tplName + Chr(10) + "Info: http://maisondesrevues.org/"
    End If
    ProcessedDoc.Range(0, ProcessedDoc.Characters.Count).Text = text
End Function

' Process général dans toutes les langues
' ==========

Private Function translateTemplate(ByVal lang As String, ByVal isComplet As Boolean)
    Dim tplName As String
    If isComplet Then
        tplName = "revuesorg_complet_" + lang + ".dot"
    Else
        tplName = "revuesorg_" + lang + ".dot"
    End If
    writeLog ""
    writeLog "Génération du modèle " + tplName
    writeLog "---------------------------------"
    copyAndOpen getAbsPath(TMPBASEDOT), getAbsPath(BUILDTEMPLATES) + "\" + tplName
    addTplInfo tplName
    translateStyles lang, isComplet
    processToolbar lang, isComplet
    saveAndClose ProcessedDoc
End Function

' TODO: destLanguage ne devrait plus être globale
Private Function processAll()
    Dim currentLang As String
    For intCount = LBound(DestLanguages) To UBound(DestLanguages)
        currentLang = Trim(DestLanguages(intCount))
        If currentLang <> "" Then
            translateTemplate currentLang, False
            translateTemplate currentLang, True
        End If
    Next
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
    Set menuItem = menuBar.Controls.Add(Type:=msoControlButton)
    With menuItem
        .Caption = "Ouvrir le répertoire de la macro"
        .OnAction = "openDestFolder"
        .Style = msoButtonCaption
    End With
End Sub

Sub openDestFolder()
    Dim path As String
    path = getAbsPath()
    If folderExists(path) Then
        Call Shell("explorer.exe" & " " & path, vbNormalFocus)
    Else
        MsgBox "Le répertoire 'revuesorg_word_templates' n'existe pas. Vérifiez que la macro est correctement installée.", vbCritical
    End If
End Sub

' Lancer la génération des modèles
Sub runGenerator()
    Dim langFound As Boolean
    Dim user As Integer
    Dim rootPath As String
    ' Vérification de l'intégrité de l'arborescence de la macro
    rootPath = getAbsPath()
    If Not folderExists(rootPath) Then
        MsgBox "Le répertoire 'revuesorg_word_templates' n'existe pas. Vérifiez que la macro est correctement installée.", vbCritical
        Exit Sub
    End If
    ' Demande de confirmation pour la fermeture des documents ouverts dans Word
    If Application.Documents.Count <> 0 Then
        user = MsgBox("L'exécution du générateur de modèles va entrainer la fermeture tous les documents actuellement ouverts dans Word. L'enregistrement sera proposé pour tous les documents qui n'ont pas été sauvegardés.", vbOKCancel + vbQuestion, "Génération des modèles")
        If user <> vbOK Then
            Exit Sub
        End If
    End If
    Call closeAll
    ' Créer les dossiers build et tmp s'ils n'existent pas déjà
    createFolder getAbsPath(TMP)
    createFolder getAbsPath(BUILD)
    createFolder getAbsPath(BUILDTEMPLATES)
    ' Convertir l'encodage des fichiers INI
    toUtf16 getAbsPath(TRANSLATIONS), getAbsPath(TMPTRANSLATIONS)
    toUtf16 getAbsPath(ENUMERATIONS), getAbsPath(TMPENUMERATIONS)
    ' Initialisation des langues
    langFound = getLanguagesFromIni()
    ' Message d'erreur et exit si les langues ne sont pas trouvées dans l'INI
    If Not langFound Then
        MsgBox "Impossible d'identifier les langues de destination. Veuillez vérifier le fichier src/translations.ini." + Chr(10) + Chr(10) + "L'exécution du générateur de modèles va être annulée.", vbCritical, "Erreur"
        Exit Sub
    End If
    Call openLog
    ' Copie et préparation de base.dot
    copyAndOpen getAbsPath(BASEDOT), getAbsPath(TMPBASEDOT)
    Call clearAllKeybindings
    saveAndClose ProcessedDoc
    ' Création d'un modele par langue déclarée : traduction des styles, génération de la toolbar, attribution des actions et des keybindings
    Call processAll
    ' Fin
    deleteFolder getAbsPath(TMP)
    Call closeLog
    Documents.Open FileName:=getAbsPath(LOGTXT), Visible:=True
    MsgBox "Les modèles ont été générés dans le dossier build/templates." + Chr(10) + Chr(10) + "Merci de vérfier qu'aucune erreur n'a été rencontrée en consultant le log des erreurs.", 64, "Opération terminée avec succès"
End Sub
