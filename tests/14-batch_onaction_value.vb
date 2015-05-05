' Test 14
' Macro utilitaire qui permet d'assigner la même valeur pour l'attribut onaction de tous les menus d'une barre d'outil.
' Voir Tests 7 et 8 pour les applications.

Const TOOLBARNAME As String = "LodelStyles"
Const MACRONAME As String = "maSuperAction"

Private Function parseSubmenuAndSetAction(submenu)
    For Each Ctl In submenu.Controls
        If Ctl.Type = msoControlButton Then
            If Ctl.HyperlinkType <> msoCommandBarButtonHyperlinkOpen Then ' Ne pas écraser les boutons de liens hypertextes
                Ctl.OnAction = MACRONAME
            End If
        ElseIf Ctl.Type = msoControlPopup Then
            parseSubmenuAndSetAction Ctl
        End If
    Next Ctl
End Function

Sub maSuperAction()
    Dim ctlCBarControl  As CommandBarControl
    Set ctlCBarControl = CommandBars.ActionControl
    If ctlCBarControl Is Nothing Then Exit Sub
    MsgBox "Vous avez cliqué sur le bouton " + ctlCBarControl.Caption
End Sub

Sub setAction()
  Dim Cmdbar As CommandBar
  Dim Ctl As CommandBarControl
  Application.ScreenUpdating = False
  For Each Cmdbar In Application.CommandBars
    If Cmdbar.Name = TOOLBARNAME Then
      For Each Ctl In Cmdbar.Controls
        parseSubmenuAndSetAction Ctl
      Next Ctl
   End If
  Next Cmdbar
End Sub
