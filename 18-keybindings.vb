' Test 18 - KeyBindings
' Attribution d'un raccourci clavier à un style ou une macro
' Enumeration des wdKey : https://msdn.microsoft.com/fr-fr/library/office/ff838929.aspx
' Enumeration des wdKeyCategory : https://msdn.microsoft.com/fr-fr/library/microsoft.office.interop.word.wdkeycategory%28v=office.11%29.aspx

Declare Function GetPrivateProfileString Lib "kernel32" Alias _
    "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
        ByVal lpKeyName As Any, ByVal lpDefault As String, _
        ByVal lpReturnedString As String, ByVal nSize As Long, _
        ByVal lpFileName As String) As Long

' Je n'ai pas trouvé le moyen d'acceder aux enumerations built in de Word. On passe donc par un ini.
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

Private Function getEnumeration(section As String, key As String)
    getEnumeration = GetSectionEntry(section, key, Options.DefaultFilePath(Path:=wdUserTemplatesPath) + "\generator\enumerations.ini")
End Function

' NOTE: La si on veut gerer toutes les combinaisons avec Ctrl, Shift et Alt, ça nous fait utiliser une série de pas moins de 8 If/ElseIf vu que BuildKeyCode() ne supporte pas les argument nuls ou non définis. En fait il semblerait que la fonction BuildKeyCode (https://msdn.microsoft.com/fr-fr/library/office/ff845364.aspx) se contente de faire la somme de ses parametres. On va donc faire pareil (une somme) en esperant tres fort que ca fonctionne a tous les coups.
Function getKeyCode(keyString As String)
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
            MsgBox "Erreur: '" + key + "' n'est pas une touche valide'"
        End If
    Next i
    getKeyCode = sum
End Function

' Fonctionne aleatoirement... il vaut mieux attribuer des raccourcis aux styles plutôt qu'aux macros
Function addMacroKeyBinding(macroName As String, paramName As String, keyString As String)
    Dim keyCode As Long
    keyCode = getKeyCode(keyString)
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
        Command:=macroName, _
        keyCode:=keyCode, _
        CommandParameter:=paramName
End Function

Function addStyleKeyBinding(styleName As String, keyString As String)
    Dim keyCode As Long
    keyCode = getKeyCode(keyString)
    ' Dans le cas d'un identifiant numerique il faut retrouver le nom du style
    If IsNumeric(styleName) Then
        styleName = ActiveDocument.Styles(CInt(styleName)).NameLocal
    End If
    Debug.Print styleName
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCategory:=wdKeyCategoryStyle, _
        Command:=styleName, _
        keyCode:=keyCode
End Function

' Une fonction pour supprimer tous les raccourcis clavier d'un template
Sub ClearAll()
    CustomizationContext = NormalTemplate
    KeyBindings.ClearAll
End Sub

' Ajouter avec le nom du style (CTRL + ALT + A)
Sub test()
    addStyleKeyBinding "Titre 2", "Control+Alt+A" ' Ce n'est pas sensible à la casse !
End Sub

' Ajouter avec l'identifiant du style natif (CTRL + ALT + A)
Sub test2()
    addStyleKeyBinding "-2", "control+Alt+A" ' -2 => Titre 1
End Sub
