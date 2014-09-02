Sub AlignNonSequentialData()

    'Copies non-sequential datasets to a new sequential data sheet.
    
    'Requires the following functions:
    '   FindLastRow
    '   FindLastCol
    '   WsExist
    '   FindMaxLongValue
    

    'Contributed to Reddit on 20140708 by random_tx_user
    
    'http://www.reddit.com/r/excel/comments/29udgm/help_creating_macro_to_fetch_data_from_a_very_big/

    Dim iTest As Long       'Counter used to iterate through tests
    Dim iTestCol As Long
    Dim iCycle As Long      'Counter used to iterate through cycles
    Dim LastCol As Long
    Dim LastRow As Long
    Dim LastCycle As Long
    Dim LastTest As Long
    Dim TestNoRow As Long
    Dim FirstDataRow As Long
    Dim TestValue As Long
    Dim HeaderRow As Long
    Dim wsCleanName As String
    Dim WsClean As Worksheet
    Dim wsOrgName As String
    Dim wsOrg As Worksheet
    Dim CycleNum As Long
    Dim i As Long
    Dim MakeItFaster As Boolean
    
    'Simple error handling
    On Error GoTo ExitSub
    
    'Set Variables / Options
    TestNoRow = 1               'Define the row where the test number will be
    HeaderRow = 2               'Define the row where the columns names will be
    FirstDataRow = 3            'Define the first row of actual data (ignoring headers, etc.)
    LastCycle = 0               'Define initial LastCycle value. Should be "0"
    wsCleanName = "CleanData"   'Set the name for the new combined data sheet
    wsOrgName = "Sheet1"        'Name of the worksheet where original data will be. "Sheet1" is default
    MakeItFaster = False        'Must be True/False. If true, turns off screen updating.
    
    If MakeItFaster Then
        Application.ScreenUpdating = False
    End If
    
    Set wsOrg = Sheets(wsOrgName)
    
    'Switch to the original data sheet
    wsOrg.Activate

    'Create a new worksheet for the clean data. Includes error checking to ensure the sheet does not
    '  already exist and option to input a different name
    If Not WsExist(wsCleanName) Then
        Worksheets.Add().Name = wsCleanName
        Set WsClean = Worksheets(wsCleanName)
    Else
        wsCleanName = InputBox("What would you like the new sheet to be named?", _
            "Worksheet " & wsCleanName & " already exists!")
            
        If wsCleanName = "" Then
            Worksheets.Add
            Set WsClean = ActiveSheet
        Else
            Worksheets.Add().Name = wsCleanName
            Set WsClean = Worksheets(wsCleanName)
        End If
    End If
    
    'Switch back to the original data sheet
    wsOrg.Activate
    
    'Find the last set of test data. Assumes each data set will have 5 columns
    LastTest = FindLastCol(TestNoRow) / 5
    
    'Find the maximum cycle number from all tests
    For iTest = 1 To LastTest
        'Defines the cycle number column for test number iTest
        iTestCol = iTest * 5 - 4
        
        'Find last row in cycle number column for test number iTest
        LastRow = FindLastRow(iTestCol)
    
        'Find max value in cycle number column
        TestValue = FindMaxLongValue(Range((Cells(FirstDataRow, iTestCol)), Cells(LastRow, iTestCol)))
            
        'If the test value is greater than stored LastCycle value, then LastCycle is updated
        If TestValue > LastCycle Then
            LastCycle = TestValue
        End If
    Next iTest
    
    'Copy the test number and data header to the new sheet
    Range(Cells(TestNoRow, 1), Cells(HeaderRow, LastTest * 5)).Copy WsClean.Cells(1, 1)
    
    WsClean.Activate
    
    'Create cycle number column that is sequential from 1 to the biggest LastCycle. Starts the count
    '  at row 3
    For iCycle = 1 To LastCycle
        Cells(iCycle + 2, 1) = iCycle
    Next iCycle

    wsOrg.Activate
    
    'Cycle through each test on wsOrg
    For iTest = 1 To LastTest
        'Defines the cycle number column for test number iTest
        iTestCol = iTest * 5 - 4
        
        'Deifines range of cycles to loop through for each each test
        LastRow = FindLastRow(iTestCol)
        For iCycle = 3 To LastRow
            Range(Cells(iCycle, iTestCol + 1), Cells(iCycle, iTestCol + 4)).Copy _
                WsClean.Cells(Cells(iCycle, iTestCol) + 2, iTestCol + 1)
        Next iCycle
    Next iTest
    
    WsClean.Activate
    
    For i = 2 To FindLastCol(2)
        If Cells(HeaderRow, i).Value = "Cycle" Then
            Columns(i).Delete
        End If
    Next i
        
'Clean up object variable and exit sub
ExitSub:
    Set wsOrg = Nothing
    Set WsClean = Nothing
    Application.ScreenUpdating = True

End Sub
