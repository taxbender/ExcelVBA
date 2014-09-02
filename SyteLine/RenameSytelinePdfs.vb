Function FindLastRow( _
    ByVal Col As Long) As Long
    
    'This function seraches for the last row in the defeind column. The
    '  asteric acts as a wildcard search so this should return any cell with
    '  a value.

    FindLastRow = Cells.Find( _
        What:="*", _
        After:=[A1], _
        SearchOrder:=xlByRows, _
        Searchdirection:=xlPrevious).Row
        
End Function

Sub OrganizePdfFiles()

Dim d As String, ext, x
Dim srcPath As String
Dim destPath As String
Dim srcFile As String

srcPath = "C:\Users\akeene\Documents\OLAttachments\"
destPath = "C:\Users\akeene\Documents\OLAttachments\Invoices\"

ext = Array("*.pdf")

    'Initialize Acrobat by creating App object
    Set gapp = CreateObject("AcroExch.App")



For Each x In ext
    d = Dir(srcPath & x)
        Do While d <> ""
        
        srcFile = srcPath & d
            
                
    
        'gapp.Show
    
        'Set AVDoc object
        Set pdfDoc = CreateObject("AcroExch.AVDoc")
        pdfDoc.Open srcFile, ""
                
        gapp.Show
        pdfDoc.BringToFront
    
    
        SendKeys "{CLEAR}"
        SendKeys "^a"
        Application.Wait Now + TimeValue("00:00:01")
        SendKeys "^c", True
        Application.Wait Now + TimeValue("00:00:01")
        SendKeys "^q"
    
        ThisWorkbook.Activate
        Sheets("Sheet1").Activate
        
    
        'Clear the sheet
        ActiveSheet.Cells.Clear
        
        Cells(1, 1).Activate
        ActiveSheet.Paste
        
        'Set the vaeriable to the invoice number in the spreadsheet.
        If Mid(d, 18, 14) = "OrderInvoicing" Then
            invnum = Left(Cells(4, 1).Value, 10)
       
        ElseIf Mid(d, 24, 17) = "ContractInvoicing" Then
            invnum = Left(Cells(2, 1).Value, 10)
            
            
        ElseIf Mid(d, 24, 12) = "SROInvoicing" Then
            invnum = Left(Cells(5, 1).Value, 10)
            
            Cells(2, 5) = invnum
        
        End If
        
                
            FileCopy srcFile, destPath & invnum & ".pdf"
            FileCopy srcFile, srcPath & "Done\" & d
            Application.Wait Now + TimeValue("00:00:01")
            Kill srcFile
            d = Dir
        Loop
Next

Set pdfDoc = Nothing
End Sub

