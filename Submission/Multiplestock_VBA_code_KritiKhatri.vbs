VBA Code for Multiple stock analysis

Sub StockAnalysis_Kriti_Khatri()


' First Loop iterates through all worksheets
    For Each ws In Worksheets
        
  ' Define and initialize variables for each ws
    Dim currentTicker As String
    
    Dim openingPrice As Double
    Dim closingPrice As Double
    
    Dim quarterlyChange As Double
    Dim percentageChange As Double
    
    Dim totalVolume As Double
    Dim summaryRow As Integer
    
    Dim lastRow As Long
    
    Dim maxVolume As Double
    Dim VolumeTicker As String
    
    Dim maxIncrease As Double
    Dim maxDecrease As Double
    Dim maxTicker As String
    Dim minTicker As String
    
        
   ' Initialize variables
    summaryRow = 2
    totalVolume = 0
    maxVolume = 0
    maxIncrease = 0
    maxDecrease = 0
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    openingPrice = ws.Cells(2, 3).Value
    
    'Setting up the summary table headers
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Quarterly Change"
    ws.Cells(1, 11).Value = "Percentage Change"
    ws.Cells(1, 12).Value = "Total Volume"
    
  'Now iterate loop through stock data from row 2 to last
   For i = 2 To lastRow
  
  'checking if the cell in new row matches the existing  or belongs to new ticker and Set current ticker symbol
   If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        currentTicker = ws.Cells(i, 1).Value
                
        ' Get closing price for the quarter
        closingPrice = ws.Cells(i, 6).Value
            
        'Percent change calculation
        quarterlyChange = closingPrice - openingPrice
        If openingPrice <> 0 Then
        percentageChange = (quarterlyChange / openingPrice)
        Else
        percentageChange = 0
        End If

        ' Add to total volume
        totalVolume = totalVolume + ws.Cells(i, 7).Value
                
        'summary results
        ws.Cells(summaryRow, 9).Value = currentTicker
        ws.Cells(summaryRow, 10).Value = quarterlyChange
        ws.Cells(summaryRow, 11).Value = percentageChange
        ws.Cells(summaryRow, 12).Value = totalVolume

        'conditional formating (green, red or default white) for quaterly change and percentage change
        If quarterlyChange > 0 Then
        ws.Cells(summaryRow, 10).Interior.Color = RGB(0, 255, 0) ' Green
        ElseIf quarterlyChange < 0 Then
        ws.Cells(summaryRow, 10).Interior.Color = RGB(255, 0, 0) ' Red
        Else
        ws.Cells(summaryRow, 10).Interior.Color = RGB(255, 255, 255) ' White
        End If
        
        If percentageChange > 0 Then
        ws.Cells(summaryRow, 11).Interior.Color = RGB(0, 255, 0) ' Green
        ElseIf percentageChange < 0 Then
        ws.Cells(summaryRow, 11).Interior.Color = RGB(255, 0, 0) ' Red
        Else
        ws.Cells(summaryRow, 11).Interior.Color = RGB(255, 255, 255) ' White
        End If

        ' Maximum Increase, decrease and Total Volume
        ' Check for highest values
        If percentageChange > maxIncrease Then
        maxIncrease = percentageChange
        maxTicker = currentTicker
        End If
 
        ' Check for lowest values
        If percentageChange < maxDecrease Then
        maxDecrease = percentageChange
        minTicker = currentTicker
        End If
        
        If totalVolume > maxVolume Then
        maxVolume = totalVolume
        VolumeTicker = tickerSymbol
        End If
        
        ' Reset variables for the next ticker
         summaryRow = summaryRow + 1
         totalVolume = 0
         openingPrice = ws.Cells(i + 1, 3).Value
         Else
         ' Add to total volume
         totalVolume = totalVolume + ws.Cells(i, 7).Value
         End If
        Next i
        
        ' Create a summary for the highest values
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Highest Volume"
        
        ' Output the highest values
        ws.Cells(2, 15).Value = maxTicker
        ws.Cells(2, 16).Value = maxIncrease
        ws.Cells(2, 16).NumberFormat = "0.00%"
        ws.Cells(3, 15).Value = minTicker
        ws.Cells(3, 16).Value = maxDecrease
        ws.Cells(3, 16).NumberFormat = "0.00%"
        ws.Cells(4, 15).Value = VolumeTicker
        ws.Cells(4, 16).Value = maxVolume
            
   ' Autofit columns for readability
        ws.Columns("I:P").AutoFit
    Next ws

End Sub



