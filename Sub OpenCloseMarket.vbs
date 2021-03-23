Sub OpenCloseMarket()

'Set the headers
Range("O1").Value = "First Open"
Range("P1").Value = "Last Close"

'Set an initial variablefor holding the stock ticker name
Dim TickerName As String

'Set an initial variable for holding the volume for each ticker name
Dim Ticker_Total As Double
Ticker_Total = 0

'Keep track of the location for each ticker name in the sumary table
Dim Summary_Ticker_Row As Integer
Summary_Ticker_Row = 2

'Set the Ticker First Open
Dim OpenTicker As Double
OpenTicker = Range("C2").Value
Range("O2") = OpenTicker

'Set the Ticker Last Close
Dim CloseTicker As Double
CloseTicker = 0

'Set the last row
LastRow = Cells(Rows.Count, 2).End(xlUp).Row

For i = 2 To LastRow

    'Check if we are still within the same ticker, if its not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        'Open Ticker Part
        
        OpenTicker = Cells(i + 1, 3).Value
        
        Cells(Summary_Ticker_Row + 1, 15).Value = OpenTicker
    
        'Close Ticker Part
        
        CloseTicker = Cells(i, 6)
    
        Range("P" & Summary_Ticker_Row).Value = CloseTicker
        
        ' Yearly Change Part
        
        Dim YearlyChange As Double
        
        YearlyChange = CloseTicker - OpenTicker
        
        Range("L" & Summary_Ticker_Row).Value = YearlyChange
        
        'Percent Change Part
        
        Dim PercentChange As Double
            
        If OpenTicker = 0 Then
            
                PercentChange = 0
                
        Else
        
                PercentChange = YearlyChange / OpenTicker
            
                Range("M" & Summary_Ticker_Row).Value = PercentChange
        
        End If
        
        ' Go to the next Series
        Summary_Ticker_Row = Summary_Ticker_Row + 1
         
    End If

Next i
    
End Sub
