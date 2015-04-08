' Test 4
' Lire un .ini
' Voir : http://vbadud.blogspot.co.uk/2008/11/how-to-read-and-write-configuration.html
' Attention : utilise la lib kernel32 qui n'est pas présente sur Mac

Public Const INIPATH = "C:\Users\t.brouard\Desktop\test.ini"

Declare Function GetPrivateProfileString Lib "kernel32" Alias _
"GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
ByVal lpKeyName As Any, ByVal lpDefault As String, _
ByVal lpReturnedString As String, ByVal nSize As Long, _
ByVal lpFileName As String) As Long

Private Function GetSectionEntry(ByVal strSectionName As String, ByVal strEntry As String, ByVal strIniPath As String) As String

    Dim X As Long
    Dim sSection As String, sEntry As String, sDefault As String
    Dim sRetBuf As String, iLenBuf As Integer, sFileName As String
    Dim sValue As String
    
    On Error GoTo ErrGetSectionentry
    sSection = strSectionName
    sEntry = strEntry
    sDefault = ""
    sRetBuf = Strings.String$(256, 0) '256 null characters
    iLenBuf = Len(sRetBuf$)
    sFileName = strIniPath
    X = GetPrivateProfileString(sSection, sEntry, _
    "", sRetBuf, iLenBuf, sFileName)
    sValue = Strings.Trim(Strings.Left$(sRetBuf, X))
    
    If sValue <> "" Then
    GetSectionEntry = sValue
    Else
    GetSectionEntry = vbNullChar
    End If
    
ErrGetSectionentry:
    If Err <> 0 Then
    Err.Clear
    Resume Next
    End If

End Function

Function getIniValue(section As String, key As String)
    Dim sIniPath
    Dim ret
    getIniValue = GetSectionEntry(section, key, INIPATH)
End Function

Sub test()
    Dim res As String
    res = getIniValue("Test", "MaVar")
    MsgBox (res)
End Sub