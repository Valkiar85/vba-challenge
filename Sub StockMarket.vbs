Sub StockMarket()

'Set the Headers
Range("K1").Value = "Ticker"
Range("L1").Value = "Yearly Change"
Range("M1").Value = "Percent Change"
Range("N1").Value = "Total Stock Volume"

'Set an initial variablefor holding the stock ticker name
Dim TickerName As String
TickerName = Cells(2, 1).Value
Range("K2").Value = TickerName

'Set an initial variable for holding the volume for each ticker name
Dim Ticker_Total As LongLong
Ticker_Total = 0

'Keep track of the location for each ticker name in the sumary table
Dim Summary_Ticker_Row As Integer
Summary_Ticker_Row = 3

'Counts the number of rows
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'Loop through each row

For i = 2 To LastRow

    'Check if we are still within the same ticker, if its not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        'Set the Ticker Name
        TickerName = Cells(i + 1, 1).Value
        
        'Add the Ticker Total
        Ticker_Total = Ticker_Total + Cells(i, 7).Value
        
        'Print the Ticker in The Summary Table
        Range("K" & Summary_Ticker_Row).Value = TickerName
        
        'Print the Ticker Amount to the Summary Table
        Range("N" & Summary_Ticker_Row - 1).Value = Ticker_Total

        ' Add one to the sumary table row
        Summary_Ticker_Row = Summary_Ticker_Row + 1
        
        'Reset the Ticker Total
        Ticker_Total = 0
        
        'If the cell immediatly following a row is the same ticker name...
        Else
        
            'Add to the Ticker Total
            Ticker_Total = Ticker_Total + Cells(i, 7).Value
            
        End If
        
    Next i

End Sub
