Attribute VB_Name = "Module1"
Sub StockMarket_summary()

'Declaring Variables
Dim ticker As String
Dim open_price As Double
Dim close_price As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim total_stock_volume As Double
Dim previous_row As Double
Dim ws As Worksheet
Dim increase_name As String
Dim decrease_name As String
Dim greatest_name As String
Dim increase As Double
Dim decrease As Double
Dim greatest_volume As Double
Dim increase_ticker As String
Dim decrease_ticker As String
Dim greatest_volume_ticker As String

'Looping all the worksheets to run the code once

For Each ws In Worksheets

'Assigning a column header for every task we are going to perform

    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
'Assigning a value to certain variables
start_data = 2
previous_row = 1
total_stock_volume = 0

'Finding the last row of the worksheet
last_row = ws.Cells(Rows.Count, "A").End(xlUp).Row

'For each ticker on each worksheet, summarize the yearly change, percent change, and total stock volume

    For i = 2 To last_row
    
    'Check for the ticker symbols if they are similar or if are different then add it under the ticker column
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    ticker = ws.Cells(i, 1).Value
    previous_row = previous_row + 1
    
    'Retriving value for the opening day and last day of the year
    
    year_open = ws.Cells(previous_row, 3).Value
    year_close = ws.Cells(i, 6).Value
    
    'Creating a for loop to add the stock volume under the same ticker name
    
    For j = previous_row To i
        total_stock_volume = total_stock_volume + ws.Cells(j, 7).Value
        
    Next j
    
    'Check if there is any 0 present in open column
    If year_open = 0 Then
        percent_change = year_close
        
    Else
        yearly_change = year_close - year_open
        percent_change = yearly_change / year_open
        
    End If
    
    
    'Putting the values under the respective headers
    ws.Cells(start_data, 9).Value = ticker
    ws.Cells(start_data, 10).Value = yearly_change
    ws.Cells(start_data, 11).Value = percent_change
    
    'Formating the Percent Change column to percentage format
    ws.Cells(start_data, 11).NumberFormat = "0.00%"
    ws.Cells(start_data, 12).Value = total_stock_volume
    
    'Foramtting the Yearly Change column
        If yearly_change > 0 Then
        ws.Cells(start_data, 10).Interior.ColorIndex = 4 'green
        Else
        ws.Cells(start_data, 10).Interior.ColorIndex = 3 'red
        End If
              
    start_data = start_data + 1
    
    'Reset the values back to 0
    total_stock_volume = 0
    yearly_change = 0
    percent_change = 0
    
    'Move i number to variable previous_row
    previous_row = i
    
    End If
    
    
Next i
Next ws
    
'
    
'Creating a second summary table that would return the stock with Greatest % increase, Greatest % decrease, and Greatest total volume.

    'Counting the number of rows in Percent Change column in each worksheet
    
    For Each ws In Worksheets
    
        klast_row = ws.Cells(Rows.Count, "K").End(xlUp).Row
        
    'Initialize variables
    
    increase = ws.Cells(2, 11).Value
    decrease = ws.Cells(2, 11).Value
    greatest_volume = ws.Cells(2, 12).Value
    increase_ticker = ws.Cells(2, 9).Value
    decrease_ticker = ws.Cells(2, 9).Value
    greatest_volume_ticker = ws.Cells(2, 9).Value
    
    'loop through each row in column K (Percent Change)
    
    For i = 2 To klast_row
        'update the greatest increase
        If ws.Cells(i, 11).Value > increase Then
            increase = ws.Cells(i, 11).Value
            increase_ticker = ws.Cells(i, 9).Value
        End If
        
        'update greatest total volume
        If ws.Cells(i, 12).Value > greatest_volume Then
            greatest_volume = ws.Cells(i, 12).Value
            greatest_volume_ticker = ws.Cells(i, 9).Value
        End If
    Next i
    
    
    'output results in the summary table
    ws.Range("N2").Value = "Greatest%Increase"
    ws.Range("N3").Value = "Greatest%Decrease"
    ws.Range("N4").Value = "GreatestTotalVolume"
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
    ws.Range("O2").Value = increase_ticker
    ws.Range("O3").Value = decrease_ticker
    ws.Range("O4").Value = greatest_volume_ticker
    ws.Range("P2").Value = increase
    ws.Range("P3").Value = decrease
    ws.Range("P4").Value = greatest_volume
    
    'format percentage for percent change values
    ws.Range("P2:P3").NumberFormat = "0.00%"
    
    Next ws
    
    
    
        
        
        







'
End Sub

