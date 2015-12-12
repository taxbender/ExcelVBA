Function getMyDocsPath() As String
  If gEnableErrorHandling Then On Error GoTo errHandler
 
  Dim oShell As Object
  Set oShell = CreateObject("WScript.Shell")
  
  getMyDocsPath = oShell.SpecialFolders("mydocuments")

exitHere:
  Set oShell = Nothing
  Exit Function

errHandler:
  Resume exitHere

End Function
