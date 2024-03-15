Attribute VB_Name = "Module1"
Sub stock_data()
                                             
'variables for the ticker name and the total volume
        Dim ticker As String
        Dim tickerV As Double
        tickerV = 0
'location for each ticker in the summary
        Dim ticksummary As Integer
        ticksummary = 2
'setting open price and headers for new columns
        Dim openingPrice As Double
        openingPrice = Cells(2, 3).Value
        
        Dim closingPrice As Double
        Dim yearlyChange As Double
        Dim percentChange As Double
        
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"

'number of rows in the first column.
'Looping through the tickers
        finalrow = Cells(Rows.Count, 1).End(xlUp).Row

        For i = 2 To finalrow

'checking if the value of the next cell changes
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
'ticker name
              ticker = Cells(i, 1).Value

'adding volume
              tickerV = tickerV + Cells(i, 7).Value

'displaying the ticker and volume
              Range("I" & ticksummary).Value = ticker
              Range("L" & ticksummary).Value = tickerV

'closing price
              closingPrice = Cells(i, 6).Value

'yearly change calculation
              yearlyChange = (closingPrice - openingPrice)
              Range("J" & ticksummary).Value = yearlyChange

                If openingPrice = 0 Then
                    percentChange = 0
                Else
                    percentChange = yearlyChange / openingPrice
                End If

'displaying yearly change in the summary
              Range("K" & ticksummary).Value = percentChange
              Range("K" & ticksummary).NumberFormat = "0.00%"
   
'Adding 1 to ticksummary
              ticksummary = ticksummary + 1

'setting volume to zero
              tickerV = 0

'opening price reset
              openingPrice = Cells(i + 1, 3)
            
            Else
              
'Adding volume
              tickerV = tickerV + Cells(i, 7).Value

            
            End If
        
        Next i

'the last row
    finalrow_Count = Cells(Rows.Count, 9).End(xlUp).Row
    
'conditional formatting
        For i = 2 To finalrow_Count
            If Cells(i, 10).Value > 0 Then
                Cells(i, 10).Interior.ColorIndex = 10
            Else
                Cells(i, 10).Interior.ColorIndex = 3
            End If
        Next i

    
'labeling the cells
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"

'Looking for the greatest and least percent change
'Looking for the maximum volume
        For i = 2 To finalrow_Count
'percent change
            If Cells(i, 11).Value = Application.WorksheetFunction.Max(Range("K2:K" & finalrow_Count)) Then
                Cells(2, 16).Value = Cells(i, 9).Value
                Cells(2, 17).Value = Cells(i, 11).Value
                Cells(2, 17).NumberFormat = "0.00%"
                
            ElseIf Cells(i, 11).Value = Application.WorksheetFunction.Min(Range("K2:K" & finalrow_Count)) Then
                Cells(3, 16).Value = Cells(i, 9).Value
                Cells(3, 17).Value = Cells(i, 11).Value
                Cells(3, 17).NumberFormat = "0.00%"
            
'volume
            ElseIf Cells(i, 12).Value = Application.WorksheetFunction.Max(Range("L2:L" & finalrow_Count)) Then
                Cells(4, 16).Value = Cells(i, 9).Value
                Cells(4, 17).Value = Cells(i, 12).Value
            
            End If
        
        Next i
        
End Sub
