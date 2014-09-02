Option Explicit

Sub CopyNonBlankCellsToIndividualLinesOnDestSheet()
        
    'Function looks for non-empty cells within a given range in one worksheet
    '   and copies them to the same cell in another worksheet.
    
    'Requires the following functions:
    '   FindLastRow
    '   FindLastCol

    'Contributed to Ozgrid on 20140520 by Aaron80
    
    On Error GoTo ExitSub         'Simple error handling
    
    'Defines all variables we plan to use in this module
    Dim srcWS As Worksheet        'Source worksheet
    Dim dstWS As Worksheet        'Destination worksheet
    Dim LastRow As Long           'Last row in given column
    Dim LastCol As Long           'Last column in given row
    Dim iRow As Long              'Interative row counter
    Dim iCol As Long              'Interative column counter
    Dim fRow As Long              'First row to start loops at
    Dim fCol As Long              'First column to start loops at
    Dim dstRow As Long            'Destination row interative counter
    Dim fdstRow As Long           'First destination row
    Dim dstCol As Long            'Column where list of contracts will be
    
    fRow = 2                        'Start loops/analysis at row 2; Row 1 is headings
    fCol = 1                        'Assumes contract data is in column 1
    Set srcWS = Sheets("Sheet1")    'Defines source contrat worksheet
    Set dstWS = Worksheets.Add      'Adds new sheet for contract copy
    dstWS.Name = "UniqueContracts"  'Rename the new sheet
    fdstRow = 2                     'First row on the destination worksheet
    dstRow = fdstRow                'Iteration counter for copy of contracts
    dstCol = 1
    
    srcWS.Activate                  'Sets the srcWS as the active sheet
    
    LastRow = FindLastRow(fCol)     'Get the last row in column fCol (currently 1)
    
    
    'First we deliminate the data into columns based on the '/' delimiter
    Columns("A:A").TextToColumns _
        Destination:=Cells(1, fCol), _
        DataType:=xlDelimited, _
        OtherChar:="/"
    
    
    'Loop through each column then row and perform a test on all cells
    '  We copy the original contract data to a new worksheet to maintain originals
    '  Depending on the size of your datathere may be more efficient ways to do
    '  this.
    
    For iRow = fRow To LastRow                   'Loop through each row
        
        'Find the last column in each row (this allows for a different number of
        '  contracts in each row
        
        LastCol = FindLastCol(iRow)
        
        For iCol = fCol To LastCol               'Loop through each column in iRow
            
            If Not Cells(iRow, iCol) = "" Then   'Test for empty cells on SrcWS
                
                'Test if cell has a value and past it into the new sheet. Everytime
                '  a new value is  pasted, add 1 to the destination row counter
                dstWS.Cells(dstRow, dstCol) = srcWS.Cells(iRow, iCol)
                
                'Iterate the destination row counter so the next row is used.
                dstRow = dstRow + 1
                
            End If
        Next iCol
    Next iRow
    

ExitSub:

    'Set our worksheet objects to nothing
    Set srcWS = Nothing
    Set dstWS = Nothing

End Sub


Function FindLastCol( _
    ByVal Row As Long) As Long

    'This function seraches for the last column in the defeind row. The
    '  asteric acts as a wildcard search so this should return any cell with
    '  a value (not sure about cells with formating but no value).
  
    FindLastCol = Cells.Find( _
        What:="*", _
        After:=[A1], _
        SearchOrder:=xlByColumns, _
        Searchdirection:=xlPrevious).Column
        
End Function


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
