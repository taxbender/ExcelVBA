Function getMyDocsPath() As String
  If gEnableErrorHandling Then On Error GoTo errHandler
  
  
  
  Dim oShell As Object
  Set oShell = CreateObject("WScript.Shell")
  
  getMyDocsPath = oShell.SpecialFolders("mydocuments")



exitHere:
  Set oShell = Nothing
  
  Exit Function

errHandler:
  MsgBox "Error " & Err.Number & ": " & Err.Description & " in ", _
          vbOKOnly, "Error"

Resume exitHere

End Function
