Sub DeleteRowsFromTheBottom()
     
    'Function  'Test if length of text in cell (i, myCol) is greater than
    '  length of text in cell (i-1, myCol). If so, cell (not row) is deleted
        
    'Requires the following functions:
    '   FindLastRow
    
    'Contributed to Ozgrid on 20140520 by Aaron80
    
    'On Error GoTo ExitSub         'Simple error handling
    
    
    Dim objWS As Worksheet          'Worksheet object changed from wsO
    Dim LastRow As Long             'Replace RowsT with LastRow
    Dim myCol As Long               'The column NUMBER you are evaluating
    Dim i As Long                   'Iterative counter
     
    
    Set objWS = Sheets("Sheet1")        'Changed OrgDatmtx to Sheet3
    myCol = 1                          'Column we are evaluating (AB = 28)
    
    objWS.Activate                      'Actives the sheet
    
    LastRow = FindLastRow(myCol)        'Finds the last row in the column
    
    For i = LastRow To 2 Step -1
        
        'Test if length of text in cell (i, myCol) is greater than
        '  length of text in cell (i-1, myCol)
        If Len(Cells(i, myCol)) > Len(Cells(i - 1, myCol)) Then
            
            'Deletes the cell (i,myCol) NOT the Row (i)
            Cells(i, myCol).Delete
        
        End If
    Next i
    

ExitSub:

    'Set our worksheet objects to nothing
    Set objWS = Nothing
     
End Sub
