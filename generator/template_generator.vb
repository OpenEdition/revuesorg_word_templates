' Déclarations

Declare Function GetPrivateProfileString Lib "kernel32" Alias _
    "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
        ByVal lpKeyName As Any, ByVal lpDefault As String, _
        ByVal lpReturnedString As String, ByVal nSize As Long, _
        ByVal lpFileName As String) As Long

Public GeneratorPath As String
Public IniPath As String
Public BasePath As String
Public BuildPath As String
Public DestLang As String
Public ProcessedDoc As Document

Private Function init()
    GeneratorPath = Options.DefaultFilePath(Path:=wdUserTemplatesPath) + "\generator"
    IniPath = GeneratorPath + "\src\translations.ini"
    BasePath = GeneratorPath + "\src\base.dot"
    BuildPath = GeneratorPath + "\build"
	DestLang = askForLang()
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

' Opérations

Private Function createNewDoc()
    Set ProcessedDoc = Documents.Add(BasePath, True, , True)
End Function

Private Function renameBaseStyles(lang As String)
	Dim newName As String
	Dim key As String
	key = lang + ".style"
	For Each style in ActiveDocument.Styles
		If style.BuiltIn = False Then ' TODO: il faut tester l'existence dans l'ini. Toutes les entrées qui n'existent pas seront préservées (index en d'autres langues par exemple) 
			newName = getIniValue(style.NameLocal, key)
			style.NameLocal = newName
		End If
	Next style
End Function

Private Function askForLang()
	' TODO: pas bon car tous les modeles doivent contenir tous les styles. Il faudra faire un for each sur toutes les langues de l'ini quand le reste sera au point
    askForLang = InputBox("Taper le code de langue de destination du modèle. La langue doit être déclérée dans le fichier translations.ini.", "Sélection de la langue")
End Function

' Exécutables

Sub testNewAndSave()
    Call init
    createNewDoc
    ProcessedDoc.Save
End Sub

Sub testNewAndRename()
    Call init
    Call createNewDoc
    renameBaseStyles(DestLang)
End Sub

Sub testLang()
	Dim inputData As String
	inputData = askForLang()
	If inputData <> "" Then
        MsgBox inputData
    End If
End Sub
