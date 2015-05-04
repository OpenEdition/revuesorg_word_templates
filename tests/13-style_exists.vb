' Test 13
' Vérifie l'existence d'un style dans le document en cours

Function styleExists(StyleName As String) As Boolean
    Dim MyStyle As Word.Style
    On Error Resume Next
    Set MyStyle = ActiveDocument.Styles(StyleName)
    styleExists = Not MyStyle Is Nothing
End Function

Sub test()
    Debug.Print styleExists("Titre") ' Vrai
    Debug.Print styleExists("Foobar") ' Faux
End Sub
