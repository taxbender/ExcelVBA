Function DiscPaybackPeriod( _
    Rate As Double, _
    CashFlows As Range) As Double

    'Function to calculate the discounted payback period
    
    'Contributed to Ozgrin on 20140717 by Aaron80
    '
    '  http://www.ozgrid.com/forum/showthread.php?t=189715
    

    Dim DCFs() As Double    'Array used to store discounted and net cash flows
    Dim i As Long           'Iteration counter
    Dim FirstRow As Long    'First row of CashFlows range
    Dim LastRow As Long     'Last row of CashFlows range
    Dim xOverPeriod As Long 'Period prior to positive net cash flow
    
    'Function is set up to handle a single column of data.
    '  If more than one column of data, then exit function; can be modified
    '  to accomodate a single row of data as well
    If CashFlows.Columns.Count <> 1 Then
        GoTo ExitFunction
    End If

    
    'Find the first and last row and column of the range
    FirstRow = CashFlows.Rows(1).Row
    LastRow = CashFlows.Rows.Count + FirstRow - 1

    
    'Resize the array to account for the range of periods; Arrays include the
    '  range (0,0) so the a 1x1 array is 2 rows and 2 columns
    ReDim DCFs(LastRow - FirstRow, 1)
    
    For i = LBound(DCFs) To UBound(DCFs)
        
        'Calculate the disocunted cash from for each period from the CashFlows range
        DCFs(i, 0) = Cells(i + FirstRow, CashFlows.Column) / ((1 + Rate) ^ i)
        
        'Calculate the net cash flow for each period. Period 0 is the outlay
        If i = 0 Then
            DCFs(i, 1) = DCFs(i, 0)
        Else
            DCFs(i, 1) = DCFs(i - 1, 1) + DCFs(i, 0)
        End If
    Next i
    
    'Find the period where net cash flow is positive
    For i = LBound(DCFs) To UBound(DCFs)
        If DCFs(i, 1) <= CashFlows(FirstRow, CashFlows.Column) Then
            xOverPeriod = i
        End If
    Next i
    
    'Payback period is the xOverPeriod plus the partial amount from the previous period
    DiscPaybackPeriod = xOverPeriod + (-DCFs(xOverPeriod, 1) / DCFs(xOverPeriod + 1, 0))
    
ExitFunction:
    Erase DCFs()        'Clears arrays values from memory
    
End Function
