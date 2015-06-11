' Test 18 - KeyBindings
' Attribution d'un raccourci clavier à un style ou une macro
' Enumeration des wdKey : https://msdn.microsoft.com/fr-fr/library/office/ff838929.aspx
' Enumeration des wdKeyCategory : https://msdn.microsoft.com/fr-fr/library/microsoft.office.interop.word.wdkeycategory%28v=office.11%29.aspx
' En fait il semblerait que la fonction BuildKeyCode https://msdn.microsoft.com/fr-fr/library/office/ff845364.aspx se contente de faire la somme de ses parametres.

' key: utiliser le num de la doc, les deux autre sont des booleens
' TODO: mettre en place des dictionnaires avec une conversion pour simplifier
Function getKeyCode(key As String, ctrl As Boolean, alt As Boolean)
    Dim keyCode As Long
    If ctrl = True And alt = True Then
        getKeyCode = BuildKeyCode(1024, 512, key)
    ElseIf ctrl = True Then
        getKeyCode = BuildKeyCode(512, key)
    ElseIf alt = True Then
        getKeyCode = BuildKeyCode(1024, key)
    Else
        getKeyCode = BuildKeyCode(key)
    End If
End Function

' Fonctionne aleatoirement... il vaut mieux attribuer des raccourcis aux styles plutôt qu'aux macros
Function addMacroKeyBinding(macroName As String, paramName As String, key As String, ctrl As Boolean, alt As Boolean)
    Dim keyCode As Long
    keyCode = getKeyCode(key, ctrl, alt)
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
        Command:=macroName, _
        keyCode:=keyCode, _
        CommandParameter:=paramName
End Function

Function addStyleKeyBinding(styleName As String, key As String, ctrl As Boolean, alt As Boolean)
    Dim keyCode As Long
    keyCode = getKeyCode(key, ctrl, alt)
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
    Dim key As Long
    addStyleKeyBinding "Titre 2", 65, True, True
End Sub


' Ajouter avec l'identifiant du style natif (CTRL + ALT + A)
Sub test2()
    Dim key As Long
    addStyleKeyBinding "-2", 65, True, True ' -2 => Titre 1
End Sub
