 Sub GetDataFromFiles()
    
        'On Error GoTo ExitMe:
        
        Dim srcDir As String
        Dim srcFileName As String
        Dim srcWsName As String
        Dim srcWS As Worksheet
        Dim destWsName As String
        Dim destWS As Worksheet
        Dim destRow As Long
        Dim FirstRow As Long
        Dim NextRow As Long
        Dim srcLastRow As Long
    
        srcDir = "C:\Users\akeene\Desktop\Sabine\SalesTaxReports\"  'Directory where forms exist; be sure to keep the trailing "\"
        srcFileName = Dir(srcDir & "*.xl??")                        'Catches .xls and .xlsx files
        srcWsName = "Sheet1"                                        'Name of form worsheet (must be this on every file)
        destWsName = "STR"                                          'Name of worksheet where summary data will be placed
        
        'Check if summary sheet exists. If it does, exit the routine and let the user fix it. If it does not, create it.
        If WsExist(destWsName) Then
            MsgBox "A worksheet named ""Summary"" already exists. Delete or rename it and try again", vbCritical
            GoTo ExitMe
        Else
            MsgBox "A worksheet named ""Summary"" does not exist. I'll create it for you.", vbInformation
            Sheets.Add.Name = destWsName
            Set destWS = ThisWorkbook.Sheets(destWsName)     'Defines the destination WS for data from forms
            
            'Create Summary Headers
            'Cells(1, 1) = "Date"
            'Cells(1, 2) = "Unit"
            'Cells(1, 3) = "Equipment"
            'Cells(1, 4) = "Worklist No"
            'Cells(1, 5) = "Activity Description"
            'Cells(1, 6) = "Phase"
            'Cells(1, 7) = ""
            'Cells(1, 8) = "Res"
            'Cells(1, 9) = "Count"
            'Cells(1, 10) = "Duration"
            'Cells(1, 11) = "Total"
            'Cells(1, 12) = "Notes"
            'Cells(1, 13) = "FileName"
        End If
        
        FirstRow = 2                                        'Default first row where summary data should be placed
        destRow = Application.Max(FindLastRow(1), FirstRow) 'Initial placement of first row of new form data
        
        
        'Loop though all files ending with the .xl** extension
        Do While srcFileName <> ""
        
            'Prevent the macro from opening the file it is saved in
            If srcFileName = ThisWorkbook.Name Then
                GoTo ExitMe
            End If
        
            'Open the form file
            Workbooks.Open (srcDir & srcFileName)
            
            'Make sure the form worksheet exists
            If Not WsExist(srcWsName) Then
                Workbooks(srcFileName).Close
                GoTo ExitMe
            Else
                'If the source WS exists set the srcWS obj
                Set srcWS = Workbooks(srcFileName).Worksheets(srcWsName)
            End If
            
            'Copy data from the src WS
            srcWS.Activate
            srcLastRow = FindLastRow(1)
            'Filters out headers and summary rows
            srcWS.Range(Cells(3, 1), Cells(srcLastRow - 2, 11)).Copy
            
            'Pate values to the destWS
            destWS.Activate
            destWS.Cells(destRow, 1).PasteSpecial xlValues
            
            'Find the new first empty row for data
            destRow = FindLastRow(1) + 1
            
            'Dump the clipboard
            Application.CutCopyMode = False
            
            'Close the source file
            Workbooks(srcFileName).Close
            
            'Get the next file (Excel magic keeps the filter from above)
            srcFileName = Dir()
            
        Loop
    
ExitMe:
    
        Application.CutCopyMode = False
        Set destWS = Nothing
        Set srcWS = Nothing
    
    End Sub
