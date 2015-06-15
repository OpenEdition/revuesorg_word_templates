' Test 19 - Prefix
' Préfixer les menus et les identifiants de styles pour éviter les interférences avec les traductions lors du traitement.

Const PREFIX As String = "$"
Const TOOLBARNAME As String = "LodelStyles"

Private Function prefixStyles()
    Dim firstChar As String
    For Each sty In ActiveDocument.Styles
        If sty.BuiltIn = False Then
            firstChar = Left(sty.NameLocal, 1)
            If firstChar <> PREFIX And firstChar <> " " Then
                Debug.Print sty.NameLocal
                sty.NameLocal = PREFIX + sty.NameLocal
                Debug.Print sty.NameLocal
            End If
        End If
    Next sty
End Function

Private Function prefixSubmenu(submenu)
    submenu.Caption = PREFIX + submenu.Caption
    ' On boucle sur les enfants du menu'
    For Each Ctl In submenu.Controls
        ' Si Ctl est un sous menu alors appel recursif de la fonction
        If Ctl.Type = msoControlPopup Then
            prefixSubmenu Ctl
        ' Si Ctl est un control...
        ElseIf Ctl.Type = msoControlButton Then
            Ctl.Caption = PREFIX + Ctl.Caption
        End If
    Next Ctl
End Function

Sub prefixAll ()
    Call prefixStyles
    Application.ScreenUpdating = False
	For Each Cmdbar In Application.CommandBars ' TODO: peut-être mieux d'utiliser la méthode .findControl() ?
	If Cmdbar.Name = TOOLBARNAME Then
		For Each Ctl In Cmdbar.Controls
            prefixSubmenu Ctl
		Next Ctl
	End If
	Next Cmdbar
End Sub
