' Test 8
' Implémentation du test 8 pour l'application d'un style

Private Sub startLodelMenu()
    Dim menuBar As CommandBar
    Dim menuItem As CommandBarControl

    Set menuBar = CommandBars.Add(menuBar:=False, Position:=msoBarTop, Name:="testbar", Temporary:=True)
    menuBar.Visible = True

    Set menuItem = menuBar.Controls.Add(Type:=msoControlButton)
    With menuItem
        .caption = "Test"
        .OnAction = "applyStyle"
        .Style = msoButtonCaption
        .tag = "Titre"
    End With
End Sub

Sub applyStyle()
    Dim ctlCBarControl  As CommandBarControl
    Dim tag As String

    Set ctlCBarControl = CommandBars.ActionControl
    If ctlCBarControl Is Nothing Then Exit Sub

    tag = ctlCBarControl.tag
    Selection.Range.Style = tag
End Sub
