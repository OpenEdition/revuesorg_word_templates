' Test 15 Write log
' La console de MSVB est limitée en nombre de lignes. On utilise un log.

Function FileExists(ByVal FileToTest As String) As Boolean
   FileExists = (Dir(FileToTest) <> "")
End Function

Sub DeleteFile(ByVal FileToDelete As String)
   If FileExists(FileToDelete) Then 'See above
      SetAttr FileToDelete, vbNormal
      Kill FileToDelete
   End If
End Sub

Sub InitLog()
    DeleteFile "C:\Users\t.brouard\Desktop\ErrorLog.txt"
    Open "C:\Users\t.brouard\Desktop\ErrorLog.txt" For Append As #1
End Sub

Sub WriteLog()
    Print #1, "Ouhouh"
End Sub

Sub CloseLog()
    Close #1
End Sub
