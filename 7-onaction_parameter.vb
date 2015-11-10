' Test 7
' Passer un paramètre à la macro qui est exécutée au moment du clic sur un bouton.
' Pour cela on retrouve l'un des attributs du boutons qui appelle la macro (voir test 7).
' Concrètement, on peut imaginer que le paramètre soit un nom de style, ce qui permettrait d'avoir une macro générique qui applique n'importe quel style (et donc de générer le menu dynamiquement dans toutes les langues)

Private Sub startLodelMenu()
    Dim menuBar As CommandBar
    Dim menuItem As CommandBarControl

    Set menuBar = CommandBars.Add(menuBar:=False, Position:=msoBarTop, Name:="testbar", Temporary:=True)
    menuBar.Visible = True

    Set menuItem = menuBar.Controls.Add(Type:=msoControlButton)
    With menuItem
        .caption = "test"
        .OnAction = "monAction"
        .Style = msoButtonCaption
        .tag = "unTagPourLetest"
        .parameter = "On peut aussi utiliser Parameter"
    End With
End Sub

Sub monAction()
    Dim ctlCBarControl  As CommandBarControl
    Dim tag As String
    Dim parameter As String

    Set ctlCBarControl = CommandBars.ActionControl
    If ctlCBarControl Is Nothing Then Exit Sub
    'Examine the Parameter property of the ActionControl to determine
    'which control has been clicked
    tag = ctlCBarControl.tag
    parameter = ctlCBarControl.parameter

    Debug.Print tag
    Debug.Print parameter
End Sub
