Attribute VB_Name = "Module1"
Sub StockMarketAnalysis()
    
    For Each ws In Worksheets
    
    ' Set variable for the last row
    Dim last_row As Long
    
    ' Set initial variable for each ticker name
    Dim ticker As String
    
    ' Set variables for opening and closing prices on the year
    Dim open_price As Double
    Dim close_price As Double
    
    open_price = ws.Cells(2, 3).Value
    
    'Keep track of each ticker name in the summary table
    Dim sum_table_row As Integer
    sum_table_row = 2
    
    ' Keep track of the sum volume for each ticker symbol
    Dim total_volume As Double
    total_volume = 0
    
    ' Set variables for the ticker with the greatest total volume and greatest increase/decrease in change
    Dim max_volume_ticker As String
    Dim max_volume As Double
    
    Dim percent_increase_ticker As String
    Dim percent_increase As Double
    
    Dim percent_decrease_ticker As String
    Dim percent_decrease As Double
    
    Dim i As Long
    
    ' Data starts from row 2, find last row of data
    last_row = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Loop through all stock market values
    For i = 2 To last_row
        
        ' Get the ticker symbol from column A
        ticker = ws.Cells(i, 1).Value
        
        ' Check if we are still within the same ticker value, if not, move to the next row
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            ' Calculate the yearly change
            close_price = ws.Cells(i, 6).Value
            ws.Cells(sum_table_row, 10).Value = close_price - open_price
            
            If ws.Cells(sum_table_row, 10).Value > 0 Then
                ws.Cells(sum_table_row, 10).Interior.ColorIndex = 4
            End If
            
            If ws.Cells(sum_table_row, 10).Value < 0 Then
                ws.Cells(sum_table_row, 10).Interior.ColorIndex = 3
            End If
            
            ' Calculate percent change
            If open_price <> 0 Then
                percent_change = ((close_price - open_price) / open_price) * 100
            Else
                percent_change = 0
            End If
            ws.Cells(sum_table_row, 11).Value = percent_change & "%"
            
            ' Print the ticker name in the summary table
            ws.Range("I" & sum_table_row).Value = ticker
        
            ' Print total volume amount in summary table
            ws.Range("L" & sum_table_row).Value = total_volume
        
            ' Check for max total volume
            If total_volume > max_volume Then
                max_volume = total_volume
                max_volume_ticker = ticker
            End If
            
            ' Check for greatest % increase
            If percent_change > percent_increase Then
                percent_increase = percent_change
                percent_increase_ticker = ticker
            End If
            
            ' Check for greatest % decrease
            If percent_change < percent_decrease Then
                percent_decrease = percent_change
                percent_decrease_ticker = ticker
            End If
            
            ' Add one to summary table row
            sum_table_row = sum_table_row + 1
        
            ' Reset total volume
            total_volume = 0
            
            ' Reset open price for next ticker
            open_price = ws.Cells(i + 1, 3).Value
            
        ' If the cell immediately following a row is the same ticker name
        Else
        
            ' Add to total volume
            total_volume = total_volume + ws.Cells(i, 7).Value
        
        End If
        
        ' Add a new column for max total volume, max increase, max decrease
        ws.Cells(4, 16).Value = max_volume_ticker
        ws.Cells(4, 17).Value = max_volume
        
        ws.Cells(2, 16).Value = percent_increase_ticker
        ws.Cells(2, 17).Value = percent_increase & "%"
        
        ws.Cells(3, 16).Value = percent_decrease_ticker
        ws.Cells(3, 17).Value = percent_decrease & "%"
        
    Next i
    
    Next ws
    
End Sub
