Attribute VB_Name = "VBA_Challenger"
Sub VBA_Challenger()

Dim ws As Worksheet
Dim lRow As Long
Dim tickerName As String
Dim i, tickerCount As Integer
Dim price As Double
Dim volume As Double
Dim openingPrice As Single
Dim closingPrice As Single
Dim rg As Range

For Each ws In Worksheets
    ws.Select
    
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    lRow = Cells(Rows.Count, 1).End(xlUp).Row
    tickerCount = 0
    volume = 0
    openingPrice = 0
    closingPrice = 0
    
    'Return Ticker, Price Change, % change and total volume
    For i = 2 To lRow
        tickerName = Cells(i, 1).Value
        If (Cells(i - 1, 1).Value <> Cells(i, 1).Value) Then
            tickerCount = tickerCount + 1
            volume = Cells(i, 7).Value
            openingPrice = Cells(i, 3).Value
        Else
            volume = volume + Cells(i, 7).Value
            closingPrice = Cells(i, 6).Value
        End If
        
        Cells(tickerCount + 1, 9).Value = tickerName
        With Cells(tickerCount + 1, 10)
            .Value = closingPrice - openingPrice
            .NumberFormat = "0.00"
        End With
        
        With Cells(tickerCount + 1, 11)
            If openingPrice = 0 Then
                .Value = 0
            Else
                .Value = (closingPrice - openingPrice) / openingPrice
            End If
                .NumberFormat = "#.##%"
        End With
        
        Cells(tickerCount + 1, 12).Value = volume
    Next i
    
    Set rg = Range("J2", Range("J2").End(xlDown))

    With rg.FormatConditions
        .Delete
        .Add(xlCellValue, xlGreater, "=0").Interior.Color = vbGreen
        .Add(xlCellValue, xlLess, "=0").Interior.Color = vbRed
    End With

    'Challenger: return greatest stats
    Dim maxIncrease, maxDecrease As Single
    Dim maxVolume As Double
    Dim maxIncreaseTicker, maxDecreaseTicker, maxVolumeTicker As String
    
    maxIncrease = Cells(2, 11).Value
    maxDecrease = Cells(2, 11).Value
    maxVolume = Cells(2, 12).Value
    maxIncreaseTicker = Cells(2, 9).Value
    maxDecreaseTicker = Cells(2, 9).Value
    maxVolumeTicker = Cells(2, 9).Value
    
    For j = 1 To tickerCount
        If Cells(j + 1, 11).Value > maxIncrease Then
            maxIncrease = Cells(j + 1, 11).Value
            maxIncreaseTicker = Cells(j + 1, 9).Value
        End If
        
        If Cells(j + 1, 11).Value < maxDecrease Then
            maxDecrease = Cells(j + 1, 11).Value
            maxDecreaseTicker = Cells(j + 1, 9).Value
        End If

        If Cells(j + 1, 12).Value > maxVolume Then
            maxVolume = Cells(j + 1, 12).Value
            maxVolumeTicker = Cells(j + 1, 9).Value
        End If
    Next j
    
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("P2").Value = maxIncreaseTicker
    Range("P3").Value = maxDecreaseTicker
    Range("P4").Value = maxVolumeTicker
    Range("Q1").Value = "Value"
    Range("Q2").Value = maxIncrease
    Range("Q2").NumberFormat = "#.##%"
    Range("Q3").Value = maxDecrease
    Range("Q3").NumberFormat = "#.##%"
    Range("Q4").Value = maxVolume
     
Next

End Sub
