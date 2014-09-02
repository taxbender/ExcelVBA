Option Explicit


Sub CompareBudgetSheets()

    'Compares two budget worksheets to identify changes. 
    '  Budgets are compared on the Delta sheet that is created.
    
    'Requires the following functions:
    '   FindLastRow
    '   WsExist

    'Contributed to Ozgrid on 20140529 by Aaron80
    
    'http://www.ozgrid.com/forum/showthread.php?t=188682&p=714903#post714903
    
    On Error GoTo ExitSub         'Simple error handling

    Dim wsBudget1 As Worksheet
    Dim wsBudget2 As Worksheet
    Dim wsDelta As Worksheet
    Dim LastRow As Long
    Dim FirstRow As Long
    Dim BudgetCol As Long
    Dim CompareWsName As String
    Dim UniqueIndexCol As Long
    Dim strFormula As String
    
'Set options here
    'First row of data on budget sheets
    FirstRow = 3
    
    'Column where amounts to compare are located
    BudgetCol = 14
    
    'Name of the worksheet where the comparison is completed
    CompareWsName = "Delta"
    
    'Column where the unique index will be located
    UniqueIndexCol = 1
        
    'Change the names as needed. I assume Budget2 is the 'newest'
    Set wsBudget1 = Worksheets("Sheet1")
    Set wsBudget2 = Worksheets("Sheet2")
    

'On to the analysis and fun stuff.

    'Create the Delta worksheet; Ask for alternate name if sheet exists
    If Not WsExist(CompareWsName) Then
        Worksheets.Add().Name = CompareWsName
        Set wsDelta = Worksheets(CompareWsName)
    Else
        CompareWsName = InputBox("What would you like the new sheet to be named?", _
            "Worksheet ""Delta"" already exists")
            
        If CompareWsName = "" Then
            Worksheets.Add
            Set wsDelta = ActiveSheet
        Else
            Worksheets.Add().Name = CompareWsName
            Set wsDelta = Worksheets(CompareWsName)
        End If
    End If
    
'Create the unique keys
    
    wsBudget1.Activate
    LastRow = FindLastRow(BudgetCol)
    
    Range(Cells(FirstRow, UniqueIndexCol), Cells(LastRow, UniqueIndexCol)).FormulaR1C1 = _
        "=RC[6]&""_""&RC[7]&""_""&RC[8]&""_""&RC[9]&""_""&rc[10]"
    
    wsBudget2.Activate
    
    LastRow = FindLastRow(BudgetCol)
    
    Range(Cells(FirstRow, UniqueIndexCol), Cells(LastRow, UniqueIndexCol)).FormulaR1C1 = _
        "=RC[6]&""_""&RC[7]&""_""&RC[8]&""_""&RC[9]&""_""&rc[10]"
    
    
'Copy keys from wsBudget1 and wsBudget2 to wsDelta, eliminate duplicates, sort accounts, and fill in the
'  values for the Budget workwheets

    wsBudget1.Activate
    LastRow = FindLastRow(BudgetCol)
    Range(Cells(FirstRow, UniqueIndexCol), Cells(LastRow, UniqueIndexCol)).Copy
    
    wsDelta.Activate
    Cells(FirstRow, UniqueIndexCol).PasteSpecial xlValues
    
    wsBudget2.Activate
    LastRow = FindLastRow(BudgetCol)
    Range(Cells(FirstRow, UniqueIndexCol), Cells(LastRow, UniqueIndexCol)).Copy
    
    wsDelta.Activate
    Cells(FindLastRow(UniqueIndexCol) + 1, UniqueIndexCol).PasteSpecial xlValues
    
    'Eliminate duplicates
    Application.CutCopyMode = False
    Range(Cells(FirstRow, UniqueIndexCol), Cells(FindLastRow(UniqueIndexCol), UniqueIndexCol)). _
        RemoveDuplicates Columns:=1, Header:=xlNo
        
    LastRow = FindLastRow(UniqueIndexCol)
        
   
    'Parse the account values from unique (Columns are based on original file)
    Range(Cells(FirstRow, UniqueIndexCol), Cells(LastRow, UniqueIndexCol)).Copy
    
    Cells(FirstRow, 7).PasteSpecial xlValues
    
    Application.CutCopyMode = False
    
    Range(Cells(FirstRow, 7), Cells(LastRow, 7)).TextToColumns _
        Destination:=Cells(FirstRow, 7), _
        DataType:=xlDelimited, _
        TextQualifier:=xlNone, _
        ConsecutiveDelimiter:=False, _
        Tab:=False, _
        Semicolon:=False, _
        Comma:=False, _
        Space:=False, _
        Other:=True, OtherChar:="_", _
        FieldInfo:=Array( _
            Array(1, 2), _
            Array(2, 2), _
            Array(3, 2), _
            Array(4, 2), _
            Array(5, 2)), _
        TrailingMinusNumbers:=True
        
    'Add column headings to Delta worksheet
    Cells(FirstRow - 1, 12) = wsBudget1.Name
    Cells(FirstRow - 1, 13) = wsBudget2.Name & " (Newest)"
    Cells(FirstRow - 1, 14) = "Delta"
    
    'Get values from wsBudget1
    wsBudget1.Activate
    LastRow = FindLastRow(UniqueIndexCol)
  
    
    strFormula = "=SUMIF(" & wsBudget1.Name & "!R" & FirstRow & "C" & UniqueIndexCol & ":R" & _
        LastRow & "C" & UniqueIndexCol & "," & wsDelta.Name & "!RC" & UniqueIndexCol & "," & _
        wsBudget1.Name & "!R" & FirstRow & "C" & BudgetCol & ":R" & _
        LastRow & "C" & BudgetCol & ")"
    
    wsDelta.Activate
    
    Range(Cells(FirstRow, 12), Cells(FindLastRow(UniqueIndexCol), 12)).Formula = strFormula
    
    'Get values from wsBudget2
    wsBudget2.Activate
    LastRow = FindLastRow(UniqueIndexCol)

    
    strFormula = "=SUMIF(" & wsBudget2.Name & "!R" & FirstRow & "C" & UniqueIndexCol & ":R" & _
        LastRow & "C" & UniqueIndexCol & "," & wsDelta.Name & "!RC" & UniqueIndexCol & "," & _
        wsBudget2.Name & "!R" & FirstRow & "C" & BudgetCol & ":R" & _
        LastRow & "C" & BudgetCol & ")"
    
    wsDelta.Activate
    
    Range(Cells(FirstRow, 13), Cells(FindLastRow(UniqueIndexCol), 13)).Formula = strFormula
    
    'Insert the Delta Formula
    Range(Cells(FirstRow, 14), Cells(FindLastRow(UniqueIndexCol), 14)).FormulaR1C1 = _
        "=Round(RC[-1]-RC[-2],2)"
        
    
    'Dang thats a lot of code for something pretty basic!
   
    GoTo ExitSub
   
ExitSub:
    Set wsBudget1 = nohting
    Set wsBudget2 = Nothing
    Set wsDelta = Nothing
       
End Sub
