Sub Stocks()

  ' Set initial variables for holding ticker name, yearly change, percent change, and total stock volume
  Dim Stock_Name As String
  Dim Yearly_Change As Double
  Dim Percent_Change As Double
  Dim Total_Stock_Volume As Double
  
  ' Set initial variables for opening and closing values for each ticker (assist with finding yearly change and percent change)
  Dim opening As Double
  Dim closing As Double
  
  ' Set an initial variable for holding the last row index for the raw data table on each worksheet
  Dim LastRow As Long
  ' Set an initial variable for holding the last row index for the newly created data table with yearly summary data for each ticker on each worksheet
  Dim LastRow_Summary_Table As Long
  
  ' Set initial variables finding the maximum, minimum, total stock volume and their respective tickers
  Dim max As Double
  Dim min As Double
  Dim max_stock_vol As Double
  Dim max_stock As String
  Dim min_stock As String
  Dim max_stock_vol_name As String
    
  ' Set initial variable to get row and increment through data in the raw data table
  Dim Sum_Table_Row As Long
  ' Set initial variable to get row and increment through data in the summary data table containing yearly change, percent change, total stock volume of each ticker
  Dim Summary_Table_Row As Long
  
  ' looping through each worksheet
  For Each ws In Worksheets
    ' get the count of rows in each worksheet to find last row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    ' initialize total counter for stock volume total, opening value, closing value, yearly change, percent change
    Total_Stock_Volume = 0
    opening = ws.Cells(2, 3).Value
    closing = ws.Cells(2, 6).Value
    Yearly_Change = 0
    Percent_Change = 0
    ' initialize variable for finding the index of the current row of the raw data table with the first row that contains data
    Sum_Table_Row = 2
    
    ' set headers for summary table for each ticker
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent_Change"
    ws.Range("L1").Value = "Total Stock Volume"
           
    ' looping through rows in the raw data table
    For i = 2 To LastRow
        ' if ticker in current row is not the ticker in the next row
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            ' add up stock volume of the same stock/ticker
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
            ' get the ticker name of the current ticker in the row
            Stock_Name = ws.Cells(i, 1).Value
            
            ' get closing value from current row
            closing = ws.Cells(i, 6).Value
            
            ' find the yearly change by subtracting closing value (for last day of the year) from opening value for (first day of the year)
            Yearly_Change = closing - opening
            ' if opening value is zero
            If opening = 0 Then
                ' percent change is 0 if opening value is 0
                Percent_Change = 0
            ' if opening value is not 0
            Else
                ' divide yearly change by opening value to find percent change
                Percent_Change = (Yearly_Change / opening)
            End If
            
            ' create rows in the new table for each ticker and its summary data with ticker name, yearly change, percent change, and total stock volume
            ws.Range("I" & Sum_Table_Row).Value = Stock_Name
            ws.Range("J" & Sum_Table_Row).Value = Yearly_Change
            ws.Range("K" & Sum_Table_Row).Value = Percent_Change
            ws.Range("L" & Sum_Table_Row).Value = Total_Stock_Volume
            
            ' format yearly change to have 2 decimal places
            ws.Range("J" & Sum_Table_Row).NumberFormat = "#,##0.00"
            
            ' if yearly change is less than 0
            If Yearly_Change < 0 Then
                ' color the yearly change cell red
                ws.Range("J" & Sum_Table_Row).Interior.ColorIndex = 3
            ' if yearly change is 0 or greater
            Else
                ' color the yearly change cell green
                ws.Range("J" & Sum_Table_Row).Interior.ColorIndex = 4
            End If

            ' format percent change to be in the percent format
            ws.Range("K" & Sum_Table_Row).NumberFormat = "0.00%"
            
            ' get the opening value of the ticker in the next row
            opening = ws.Cells(i + 1, 3).Value
            ' set the stock volume total to 0
            Total_Stock_Volume = 0
            ' increment row value by 1 (to keep track of the row loop iterating through)
            Sum_Table_Row = Sum_Table_Row + 1
        ' if ticker in current row is the same as the ticker in the next row
        Else
            ' add up stock total volume of the current row to current stock total volume
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
        End If
    Next i
  
    ' get the row count of the newly created summary table for stocks/tickers for the year
    LastRow_Summary_Table = ws.Cells(Rows.Count, 11).End(xlUp).Row
    ' set greatest increase in percentage (max), greatest decrease in percentage (min), greatest total stock volume (max_stock_vol) and their corresponding ticker value names to that of the corresponding values in row 2
    max = ws.Cells(2, 11).Value
    min = ws.Cells(2, 11).Value
    max_stock = ws.Cells(2, 9).Value
    min_stock = ws.Cells(2, 9).Value
    max_stock_vol = ws.Cells(2, 12).Value
    max_stock_vol_name = ws.Cells(2, 9).Value
    
    For j = 2 To LastRow_Summary_Table ' loop through rows in the newly created summary table for stocks/tickers for the year
        If ws.Cells(j, 11).Value > max Then ' if the current percent change is larger than the old max
            max = ws.Cells(j, 11).Value ' store it as the new max
            max_stock = ws.Cells(j, 9).Value ' get the corresponding ticker name and store it
        End If
        If ws.Cells(j, 11).Value < min Then ' if the current percent change is smaller than the old min
            min = ws.Cells(j, 11).Value ' store it as the new min
            min_stock = ws.Cells(j, 9).Value ' get the corresponding ticker name and store it
        End If
        If ws.Cells(j, 12).Value > max_stock_vol Then ' if the current total stock volume is larger than the old max_stock_vol
            max_stock_vol = ws.Cells(j, 12).Value ' store it as the new max_stock_vol
            max_stock_vol_name = ws.Cells(j, 9).Value ' get the corresponding ticker name and store it
        End If
    Next j
    
    ' set up headers for columns and rows in a new table containing extreme values for the years
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    ' fill the table with greatest increase in percentage, greatest decrease in percentage, and the greatest total stock volume and their corresponding ticker names in the corresponding cells
    ws.Range("P2").Value = max_stock
    ws.Range("Q2").Value = max
    ws.Range("P3").Value = min_stock
    ws.Range("Q3").Value = min
    ws.Range("P4").Value = max_stock_vol_name
    ws.Range("Q4").Value = max_stock_vol
    
    ' format cells with values for the greatest increase in percentage and greatest decrease in percentage to be in percent format
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3").NumberFormat = "0.00%"
        
  Next ws
  
End Sub