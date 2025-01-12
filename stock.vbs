Sub LoopThroughStocks()
    ' variables
    Dim tickerSymbol As String
    Dim lastRow As Long
    Dim currentRow As Long
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim quarterlyChange As Double
    Dim percentageChange As Double
    Dim totalVolume As Double
    Dim outputRow As Long
    
    Dim ws As Worksheet
    Set ws = ActiveSheet

    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Start output from a new column (e.g., starting at column E)
    outputRow = 1
    ws.Cells(outputRow, 8).Value = "Ticker"
    ws.Cells(outputRow, 9).Value = "Quarterly Change"
    ws.Cells(outputRow, 10).Value = "Percentage Change"
    ws.Cells(outputRow, 11).Value = "Total Volume"
    outputRow = outputRow + 1

    ' Loop through each row of data
    For currentRow = 2 To lastRow ' Assuming row 1 contains headers
        ' Get the ticker symbol from column A
        tickerSymbol = ws.Cells(currentRow, 1).Value

        ' Get the opening price from column B
        openingPrice = ws.Cells(currentRow, 3).Value

        ' Get the closing price from column C
        closingPrice = ws.Cells(currentRow, 6).Value

        ' Get the total volume from column D
        totalVolume = ws.Cells(currentRow, 7).Value

        ' Calculate quarterly change
        quarterlyChange = closingPrice - openingPrice
        
' Calculate percentage change
        percentageChange = (quarterlyChange / openingPrice) * 100

        ' Output the results to the worksheet
        ws.Cells(outputRow, 8).Value = tickerSymbol
        ws.Cells(outputRow, 9).Value = quarterlyChange
        ws.Cells(outputRow, 10).Value = Format(percentageChange, "0.00") & "%"
        ws.Cells(outputRow, 11).Value = totalVolume

        outputRow = outputRow + 1
    Next currentRow

   ' Track greatest percentage increase
        If percentageChange > greatestIncrease Then
            greatestIncrease = percentageChange
            greatestIncreaseTicker = tickerSymbol
        End If

        ' Track greatest percentage decrease
        If percentageChange < greatestDecrease Then
            greatestDecrease = percentageChange
            greatestDecreaseTicker = tickerSymbol
        End If

        ' Track greatest total volume
        If totalVolume > greatestVolume Then
            greatestVolume = totalVolume
            greatestVolumeTicker = tickerSymbol
        End If
          ' Output greatest values to the worksheet
    ws.Cells(1, 13).Value = "Greatest % Increase"
    ws.Cells(2, 13).Value = "Greatest % Decrease"
    ws.Cells(3, 13).Value = "Greatest Total Volume"

    ws.Cells(1, 14).Value = greatestIncreaseTicker
    ws.Cells(2, 14).Value = greatestDecreaseTicker
    ws.Cells(3, 14).Value = greatestVolumeTicker

    ws.Cells(1, 15).Value = Format(greatestIncrease, "0.00") & "%"
    ws.Cells(2, 15).Value = Format(greatestDecrease, "0.00") & "%"
    ws.Cells(3, 15).Value = greatestVolume
    
      MsgBox "Quarterly data processed. Check the output columns."
End Sub

