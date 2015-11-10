' Test 17 - Supprimer les styles "Car"
' Word ajoute parfois des styles "datepublioeuvre Car Car". Il y en a plusieurs cachés dans les modèles de Revues.org
' Pour les supprimer il faut passer par une manipulation bizarre. Voir http://homepage.swissonline.ch/cindymeister/MyFavTip.htm#CharStyl
' ...mais je ne parviens pas à tous les supprimer (restent " Car Car", " Car Car1", " Car Car2")

Function deleteChar(styleToDel As Style)
    Dim styl As Style
    Set styl = ActiveDocument.Styles.Add(Name:="Style1")
    On Error Resume Next
    styleToDel.LinkStyle = styl
    styl.Delete
End Function

Function cleanStyles()
	Dim sty As Style
	For Each sty In ActiveDocument.Styles
		If sty.BuiltIn = False And sty.NameLocal Like "* Car*" Then
			deleteChar sty
		End If
	Next sty
End Function

Sub listStyles()
    For Each sty In ActiveDocument.Styles
    If sty.BuiltIn = False Then
        Debug.Print sty.NameLocal
    End If
    Next sty
End Sub

Sub test()
    Call cleanStyles
End Sub