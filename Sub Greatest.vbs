Sub Greatest()

'Headers

Range("T1").Value = "Ticker"
Range("U1").Value = "Value"
Range("S2").Value = "Greatest Percentage Increase"
Range("S3").Value = "Greatest Percentage Decrease"
Range("S4").Value = "Greatest Total Volume"

Dim TickerIncrease As String
Dim TickerDecrease As String
Dim TickerVolume As String

Dim MaxTicker, MinTicker As Double
Dim MaxVolume As LongLong

MaxTicker = 0
MinTicker = 0
MaxVolume = 0

'Counts the number of rows
LastRow = Cells(Rows.Count, 11).End(xlUp).Row

For i = 2 To LastRow

    If Cells(i, 13).Value > MaxTicker Then
    
    MaxTicker = Cells(i, 13).Value
    TickerIncrease = Cells(i, 11).Value
    
    Range("T2").Value = TickerIncrease
    Range("U2").Value = MaxTicker
    Range("U2").NumberFormat = "0.00%"
    
    End If
    
    If Cells(i, 13).Value < MinTicker Then
    
    MinTicker = Cells(i, 13).Value
    TickerDecrease = Cells(i, 11).Value
    
    Range("T3").Value = TickerDecrease
    Range("U3").Value = MinTicker
    Range("U3").NumberFormat = "0.00%"
    
    End If
    
    If Cells(i, 14).Value > MaxVolume Then
    
    MaxVolume = Cells(i, 14).Value
    TickerVolume = Cells(i, 11).Value
    
    Range("T4").Value = TickerVolume
    Range("U4").Value = MaxVolume

    End If

Next i

End Sub
