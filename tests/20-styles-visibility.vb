' Test 20 - Visibilité des styles
' Définir la visibilité des styles :  styles recommandés dans le panneau de styles (Style.Visibility) et styles rapides (Style.QuickStyle)
' FIXME: ne fonctionne pas :/ Les propriétés sont correctement modifiées pourtant l'interface n'est pas mise à jour...

Const TOOLBARNAME As String = "LodelStyles"

Private Function hideStyles()
    Dim Item As style
    For Each Item In ActiveDocument.Styles
        Item.Visibility = False
        If Item.QuickStyle <> False Then
            Item.QuickStyle = False
        End If
    Next
End Function

Private Function setVisible(submenu)
    Dim styleName As String
    ' On boucle sur les enfants du menu'
    For Each Ctl In submenu.Controls
        If Ctl.Type = msoControlPopup Then
            setVisible Ctl
        ' Si Ctl est un control...
        ElseIf Ctl.Type = msoControlButton Then
            styleName = Ctl.parameter
            If styleName <> "" Then
                If IsNumeric(styleName) Then
                    styleName = CInt(styleName)
                End If
                With ActiveDocument.Styles(styleName)
                    Debug.Print .NameLocal
                    .Visibility = True
                    .QuickStyle = True
                    Debug.Print .Visibility
                    Debug.Print .QuickStyle
                End With
            End If
        End If
    Next Ctl
End Function

Sub test()
    Call hideStyles
    For Each Cmdbar In Application.CommandBars ' TODO: peut-être mieux d'utiliser la méthode .findControl() ?
        If Cmdbar.Name = TOOLBARNAME Then
            For Each Ctl In Cmdbar.Controls
                setVisible Ctl
            Next Ctl
        End If
    Next Cmdbar
    ActiveDocument.UpdateStyles
End Sub
