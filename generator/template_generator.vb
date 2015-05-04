' Déclarations

' NOTE: translations.ini doit être encodé en UTF16-LE ou ANSI
Declare Function GetPrivateProfileString Lib "kernel32" Alias _
    "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
        ByVal lpKeyName As Any, ByVal lpDefault As String, _
        ByVal lpReturnedString As String, ByVal nSize As Long, _
        ByVal lpFileName As String) As Long

Public GeneratorPath As String
Public IniPath As String
Public BasePath As String
Public BuildPath As String
Public TmpPath As String
' Public DestLang As String ' FIXME: obsolete
Public DestLanguages() As String
Public ProcessedDoc As Document

Private Function init()
    GeneratorPath = Options.DefaultFilePath(Path:=wdUserTemplatesPath) + "\generator"
    IniPath = GeneratorPath + "\src\translations.ini"
    BasePath = GeneratorPath + "\src\base.dot"
    TmpPath = GeneratorPath + "\tmp"
    BuildPath = GeneratorPath + "\build"
    ' DestLang = askForLang() ' FIXME: obsolete
    Call getLanguagesFromIni
End Function

' FIXME: Obsolete. Maintenant on boucle sur toutes les langues
Private Function askForLang()
    Dim inputData As String
    inputData = InputBox("Taper le code de langue de destination du modèle. La langue doit être déclérée dans le fichier translations.ini.", "Sélection de la langue")
    If inputData <> "" Then ' TODO: tester aussi l'existence de la langue dans l'INI
        askForLang = inputData
    End If
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

Private Function getIniValue(section As String, key As String)
    Dim sIniPath
    Dim ret
    getIniValue = GetSectionEntry(section, key, IniPath)
End Function

' File operations

Private Function copyAndOpen(src As String, dest As String)
    FileCopy src, dest
    Set ProcessedDoc = Documents.Open(FileName:=dest, Visible:=True) ' TODO: Visible := False
End Function

Private Function saveAndClose(doc As Document)
    doc.Save
    doc.Close
End Function

' Traduction des styles
' 1. Renommer les styles déjà présents dans base.dot d'après le fichier INI = BaseStyle. C'est l'entrée "style" sans prefixe de chaque section qui est utilisée (=français).
' 2. Pour chaque langue, regarder si une traduction existe dans le INI. Si elle existe et aucun style à son nom n'existe alors on le créée. Le style créé est hérité du BaseStyle.
Private Function renameBaseStylesAndTranslate()
    Dim styleId As String
    Dim newName As String
    Dim intCount As Integer
    Dim currentLang As String

    For Each Style In ActiveDocument.Styles ' FIXME: ActiveDocument ok ?
        If Style.BuiltIn = False Then
            ' Renommer le BaseStyle
            styleId = Style.NameLocal
            newName = getIniValue(styleId, "style")
            If newName <> vbNullChar Then
                Debug.Print "newName " + newName
                Style.NameLocal = newName
            End If
            ' Le dupliquer en autant de traductions que nécessaire
            For intCount = LBound(DestLanguages) To UBound(DestLanguages)
                currentLang = Trim(DestLanguages(intCount))
                translateStyle Style.NameLocal, styleId, currentLang
            Next
        End If
    Next Style
End Function

Private Function translateStyle(baseStyleName As String, styleId As String, lang As String)
    Dim translatedName As String
    Dim key As String
    Dim styleAdded As Style

    key = lang + ".style"
    translatedName = getIniValue(styleId, key)
    If translatedName = vbNullChar Or styleExists(translatedName) Then
        Exit Function
    Else
        ' FIXME: ActiveDocument ok ?
        Set styleAdded = ActiveDocument.Styles.Add(Name:=translatedName, _
            Type:=wdStyleTypeParagraph)
        styleAdded.baseStyle = baseStyleName
    End If
End Function

Private Function styleExists(StyleName As String) As Boolean
    Dim MyStyle As Word.Style
    On Error Resume Next
    Set MyStyle = ActiveDocument.Styles(StyleName) ' FIXME: ActiveDocument ok ?
    styleExists = Not MyStyle Is Nothing
End Function

' Si nécessaire (traduction différente) copier les styles de base et traduire la copie
Private Function translateStyles(lang As String)
    Dim newName As String
    Dim key As String
    key = lang + ".style"
    For Each Style In ActiveDocument.Styles
        If Style.BuiltIn = False Then ' TODO: il faut tester l'existence dans l'ini. Toutes les entrées qui n'existent pas seront préservées (index en d'autres langues par exemple)
            newName = getIniValue(Style.NameLocal, key)
            Style.NameLocal = newName
        End If
    Next Style
End Function

' Tests
Sub test()
    Call init
    copyAndOpen BasePath, TmpPath + "\test.dot"
    Call renameBaseStylesAndTranslate
    saveAndClose ProcessedDoc
End Sub
