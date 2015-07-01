' Test 21 - Lire un INI en Unicode
' VBA gère l'encodage de ses Strings n'importe comment. Voir : http://blog.nkadesign.com/2013/vba-unicode-strings-and-the-windows-api/
' Ce test est une version améliorée du Test 4

' Contenu du fichier ini de test :
'   [tests]
'   fr="Hééééààöo !"
'   ar="ملخص، خلاصة (ar)"
'   el="περίληψη (el)"
'   he="(he) תקציר"
'   mk="македонски јазик"
'   ja="日本語"

' Lire INI avec support d'Unicode. Le fichier source doit être encodé en UTF-16 LE.
' Voir : http://www.access-programmers.co.uk/forums/showthread.php?t=164136
Public Declare Function GetPrivateProfileString _
                            Lib "Kernel32" Alias "GetPrivateProfileStringW" _
                            (ByVal lpApplicationName As String, _
                             ByVal lpKeyName As String, _
                             ByVal lpDefault As String, _
                             lpReturnedString As Any, _
                             ByVal nSize As Long, _
                             ByVal lpFileName As String) As Long

Public Function ProfileGetItem(sSection As String, _
                                sKeyName As String, _
                                sInifile As String) As String
    Dim retval As Long
    Dim cSize As Long
    Dim Buf() As Byte
    ReDim Buf(254)
    cSize = 255
    Dim sDefValue As String
    sDefValue = ""
    retval = GetPrivateProfileString(StrConv(sSection, vbUnicode), _
                                        StrConv(sKeyName, vbUnicode), _
                                        StrConv(sDefValue, vbUnicode), _
                                        Buf(0), _
                                        cSize, _
                                        StrConv(sInifile, vbUnicode))
    If retval > 0 Then
        ProfileGetItem = Left(Buf, retval)
    Else
        ProfileGetItem = vbNullChar
    End If
End Function

' Dans cet exemple on créée un menu avec la chaîne obtenue
Private Function createTestMenu(caption As String)
    If caption = vbNullChar Then Exit Function
    Dim menuBar As CommandBar
    Dim menuItem As CommandBarControl
    Set menuBar = CommandBars.Add(menuBar:=False, Position:=msoBarTop, Name:="Test21", Temporary:=True)
    menuBar.Visible = True
    Set menuItem = menuBar.Controls.Add(Type:=msoControlButton)
    With menuItem
        .caption = caption
        .Style = msoButtonCaption
    End With
End Function

Sub test()
    Dim testString As String
    testString = ProfileGetItem("tests", "ar", Options.DefaultFilePath(Path:=wdUserTemplatesPath) + "\test-encodage-utf16.ini")
    createTestMenu testString
End Sub
