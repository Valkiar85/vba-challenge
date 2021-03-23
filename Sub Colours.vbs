Sub Colours():

'Counts the number of rows
LastRow = Cells(Rows.Count, 11).End(xlUp).Row

For i = 2 To LastRow

    If Cells(i, 13).Value >= 0 Then
        Cells(i, 13).Interior.ColorIndex = 4
        
    Else
        Cells(i, 13).Interior.ColorIndex = 3
        
    End If

Next i

End Sub
