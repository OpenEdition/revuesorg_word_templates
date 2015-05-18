' Test 16 - Conversion Unicode to ANSI
' Convertir le fichier INI en ANSI afin qu'il soit lisible par Word.
' Comme ça on n'a pas à utiliser des encodages ANSI ou UTF-16 lors de la rédaction du INI.

Sub unicode2ansi(source As String, dest As String)
    Dim strText
    With CreateObject("ADODB.Stream")
        .Type = 2 'Specify stream type - we want To save text/string data.
        .Charset = "utf-8" 'Specify charset For the source text data.
        .Open 'Open the stream And write binary data To the object
        .LoadFromFile source
        strText = .ReadText(-1)
        .Position = 0
        .SetEOS
        .Charset = "_autodetect_all"
        .WriteText strText, 0
        .SaveToFile dest, 2
        .Close
    End With
End Sub

Sub test()
    unicode2ansi "C:\Users\t.brouard\Desktop\translations-utf8.ini", "C:\Users\t.brouard\Desktop\ansi.ini"
End Sub