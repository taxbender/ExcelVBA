Private Function TestURL(strUrl As String) As Boolean
    
    'Requires Microsoft WinHTTP Services, Version 5.1
    
    'Function returns TRUE is URL is found; FALSE if not found
    
    Dim oURL As New WinHttpRequest

    On Error GoTo TestURL_Err
     
    With oURL
        .Open "GET", strUrl, False
        .send
        If .Status = 200 Then          '200 indicates resource was retrieved
            TestURL = True
        End If
      TestURL = (.Status = 200)
    End With
     
TestURL_Err:
    Set oURL = Nothing                 'Clean up object
    TestURL = False                    'Invalid resources cause an error

End Function
