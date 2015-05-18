' Macro d'application de style
' Cette macro doit impérativement être présente dans base.dot
Sub ApplyLodelStyle()
    Dim ctlCBarControl  As CommandBarControl
    Dim parameter As String

    Set ctlCBarControl = CommandBars.ActionControl
    If ctlCBarControl Is Nothing Then Exit Sub

    parameter = ctlCBarControl.parameter
    If ctlCBarControl.parameter <> "" Then
        Selection.Range.Style = parameter
        Debug.Print "Application de " + parameter ' TODO: supprimer ce message
    Else
        Debug.Print "Le menu ne comporte pas de parametre."
    End If
End Sub
