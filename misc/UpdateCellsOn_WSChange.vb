Private Sub Worksheet_Change(ByVal Target As Range)

    Dim FirstRow As Long
    Dim LastRow As Long
    Dim FirstCol As Long
    Dim LastCol As Long
    Dim i As Long

    Dim ColForUserName As Long
    Dim ColForDate As Long
    Dim ColForTime As Long

    'Allows the macro to run once rather than multiples times while
    '  changing cell values
    Application.EnableEvents = False

    'Option to set location of the username, date, and time stamps
    ColForUserName = 14
    ColForDate = 15
    ColForTime = 16

    'Find the first and last row and column of the range
    FirstRow = Target.Rows(1).Row
    LastRow = Target.Rows.Count + FirstRow - 1
    FirstCol = Target.Columns(1).Column
    LastCol = Target.Columns.Count + FirstCol - 1

     'Exit the macro if anything after Colum L is changed.
    If LastCol > 12 Then
        GoTo ExitSub
    End If

    'Add username, date, and time when data in the row was changed
    '  we use a loop to allow users to change multiple rows at once
    For i = FirstRow To LastRow
        Cells(i, ColForUserName) = Environ$("Username")
        Cells(i, ColForDate) = Date
        Cells(i, ColForTime) = Time
    Next i

    'Analyze the value in column E for each row
    For i = FirstRow To LastRow
        If Cells(i, 5) = "Plunger Lift" Then
        'Reference your macro here

        'If Cell in column E is anything other than "Plunger Lift" clear the cells
        '  in column G to M
        Else
            Range(Cells(i, 7), Cells(i, 13)).Clear
            Cells(i, 10).Interior.ColorIndex = -4142
            Cells(i, 13).Interior.ColorIndex = -4142
        End If
    Next i
    'Puts setting back to default and exits the sub
ExitSub:

    Application.EnableEvents = True

End Sub
