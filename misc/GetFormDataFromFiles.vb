Sub GetFormDataFromFiles()

    'Open standardized 'form' files in the given directory and copies the data
    'to a spreadsheet in a uniform table fashion

    'Requires the following Functions
    '  FindLastRow
    '  WSExist
    
    'Contributed to Reddit on 20140623 by randon_tx_user
    '  Link http://www.reddit.com/r/excel/comments/28spp7/vba_pulling_cell_data_into_summary_to_prep_for/
    
    On Error GoTo ExitMe:
    
    Dim srcDir As String
    Dim srcFileName As String
    Dim srcWsName As String
    Dim srcWS As Worksheet
    Dim destWsName As String
    Dim destWS As Worksheet
    Dim destRow As Long
    Dim FirstRow As Long
    Dim NextRow As Long
    
    'Dim sn2 As String
    'Dim n As Long
    'NR As Long


    srcDir = "C:\Test\"                         'Directory where forms exist; be sure to keep the trailing "\"
    srcFileName = Dir(srcDir & "*.xl??")            'Catches .xls and .xlsx files
    srcWsName = "Area"                              'Name of form worsheet (must be this on every file)
    destWsName = "Summary"                          'Name of worksheet where summary data will be placed
    
    'Check if summary sheet exists. If it does, exit the routine and let the user fix it. If it does not, create it.
    If WsExist(destWsName) Then
        MsgBox "A worksheet named ""Summary"" already exists. Delete or rename it and try again", vbCritical
        GoTo ExitMe
    Else
        MsgBox "A worksheet named ""Summary"" does not exist. I'll create it for you.", vbInformation
        Sheets.Add.Name = destWsName
        Set destWS = ThisWorkbook.Sheets(destWsName)     'Defines the destination WS for data from forms
        
        'Create Summary Headers
        Cells(1, 1) = "Date"
        Cells(1, 2) = "Unit"
        Cells(1, 3) = "Equipment"
        Cells(1, 4) = "Worklist No"
        Cells(1, 5) = "Activity Description"
        Cells(1, 6) = "Phase"
        Cells(1, 7) = ""
        Cells(1, 8) = "Res"
        Cells(1, 9) = "Count"
        Cells(1, 10) = "Duration"
        Cells(1, 11) = "Total"
        Cells(1, 12) = "Notes"
        Cells(1, 13) = "FileName"
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
        End If
        
        'Set the source WS variable to facilitate copying data from form file
        Set srcWS = Workbooks(srcFileName).Worksheets(srcWsName)
        
        'Copy and paste value for lines 17 to 38 of form
        srcWS.Range(Cells(17, 1), Cells(38, 9)).Copy
        destWS.Activate
        destWS.Cells(destRow, 5).PasteSpecial xlValues
        
        
        'Copy form header information to destination table
        
        'Date
        destWS.Range(Cells(destRow, 1), Cells(destRow + 21, 1)) = srcWS.Cells(3, 3).Value
        
        'Unit
        destWS.Range(Cells(destRow, 2), Cells(destRow + 21, 2)) = srcWS.Cells(7, 3).Value

        'Equipment
        destWS.Range(Cells(destRow, 3), Cells(destRow + 21, 3)) = srcWS.Cells(15, 3).Value

        'Worklist No
        destWS.Range(Cells(destRow, 4), Cells(destRow + 21, 4)) = srcWS.Cells(9, 3).Value

        'Filename
        destWS.Range(Cells(destRow, 14), Cells(destRow + 21, 14)) = srcDir & srcFileName
        
        Application.CutCopyMode = False
        
        'Close the form workbook
        Workbooks(srcFileName).Close

        
        'Find the new lastrow on the dummary sheet
        NextRow = FindLastRow(1) + 1
        
        'Goto the next form file
        srcFileName = Dir()
    Loop
        

ExitMe:

    Application.CutCopyMode = False
    Set destWS = Nothing
    Set srcWS = Nothing

End Sub
