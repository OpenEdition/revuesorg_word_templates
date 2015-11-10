' Test 16 - Conversion UTF-8 to UTF-16
' Convertir le fichier INI en UTF-16 afin qu'il soit lisible par Word.
' Comme ca on n'a pas a utiliser des encodages ANSI ou UTF-16 lors de la redaction du INI.

Sub toUtf16(source As String, dest As String)
    Dim strText
    With CreateObject("ADODB.Stream")
        .Type = 2 'Specify stream type - we want To save text/string data.
        .Charset = "utf-8" 'Specify charset For the source text data.
        .Open 'Open the stream And write binary data To the object
        .LoadFromFile source
        strText = .ReadText(-1)
        .Position = 0
        .SetEOS
        .Charset = "utf-16"
        .WriteText strText, 0
        .SaveToFile dest, 2
        .Close
    End With
End Sub

Sub test()
    toUtf16 "C:\Users\t.brouard\Desktop\utf8.ini", "C:\Users\t.brouard\Desktop\utf16.ini"
End Sub
