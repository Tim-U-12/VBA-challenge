Attribute VB_Name = "Module1"
' Loops through all the stocks for one year and outputs the following information:
' Done - Ticker Symbol
' Done - Yearly Change from the opening price at the beginning of a give year to the closing price at the end of the year
' Done - Percentage change from the opening price at the beginning of a given year to the closing price at the end of the that year
' Done - The total stock volume of the stock
' ------------------------------------------------------------------------------------------------------------------------------------
' Bonus (+ Ticker|Name):
' Done - Greatest % increase
' Done - Greatest % decrease
' Done - Greatest Total Volume


Sub Stock_Analysis():
    ' Loops through all of the worksheets
    For Each ws In Worksheets
        Dim WorksheeName As String
        WorksheetName = ws.Name
        
        ' loop index
        Dim i As Integer
        
        ' end index
        Dim end_i As Integer
        end_i = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' init year start open
        Dim yr_start_open As Double
        yr_start_open = ws.Cells(2, 3).Value
        
        ' Yearly Change
        Dim yr_change As Double
        
        ' Yearly Percentage Change
        Dim yr_percent_change As Double
        
        ' --------------------
        '        BONUS
        ' --------------------
        ' Greatest Percentage Increase
        greatest_percentage_increase = Null
        Dim greatest_percentage_increase_ticker As String
        
        ' Greatest Percentage Change
        greatest_percentage_decrease = Null
        Dim greatest_percentage_decrease_ticker As String
        
        ' Greatest Total Volume
        greatest_total_vol = Null
        Dim greatest_total_vol_ticker As String
        
        

        ' #############################################################################################
        
        ' Titles for summary Table + inserting new column
        ws.Range("i1").EntireColumn.Insert
        ws.Range("i1").Value = "Ticker"
        
        ws.Range("j1").EntireColumn.Insert
        ws.Range("j1").Value = "Yearly Change"
        
        ws.Range("k1").EntireColumn.Insert
        ws.Range("k1").Value = "Percentage Change"
        
        ws.Range("l1").EntireColumn.Insert
        ws.Range("l1").Value = "Total Stock Volume"
        
        ws.Range("N1").EntireColumn.Insert
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        
        ws.Range("O1").EntireColumn.Insert
        ws.Range("O1").Value = "Ticker"
        
        ws.Range("P1").EntireColumn.Insert
        ws.Range("P1").Value = "Value"
        
        ' Change width of summary columns
        ws.Columns("J:L").ColumnWidth = 20
        
        ' Change width of Bonus cells
        ws.Columns("N:P").ColumnWidth = 20
        
        ' Change the format of column k to be in percentages
        ws.Columns("k").NumberFormat = "0.00%"
        
        ws.Range("P2", "P3").NumberFormat = "0.00%"
        
        ' #############################################################################################
        
        stock_volume = 0
        ' Iterate through the rows in a single worksheet
        For i = 2 To end_i
            stock_volume = stock_volume + ws.Cells(i, 7).Value
        
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                ' Adding unique ticker to summary table
                ticker = ws.Cells(Rows.Count, 9).End(xlUp).Row + 1
                ws.Cells(ticker, 9).Value = ws.Cells(i, 1).Value
                
                ' Adding the difference between open and close to the summary table
                yr_change = ws.Cells(i, 6).Value - yr_start_open
                yr_change_index = ws.Cells(Rows.Count, 10).End(xlUp).Row + 1
                ws.Cells(yr_change_index, 10).Value = yr_change
                
                ' Format the yearly change outcomes
                If yr_change < 0 Then
                    ws.Cells(yr_change_index, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(yr_change_index, 10).Interior.ColorIndex = 4
                End If
                
                ' Percentage change from the opening price at the beginning of a given year to end
                yr_percent_change = ws.Cells(i, 6).Value / yr_start_open
                yr_percent_change_index = ws.Cells(Rows.Count, 11).End(xlUp).Row + 1
                ws.Cells(yr_percent_change_index, 11).Value = yr_percent_change - 1
                                
                ' Adding the volume stock to the summary table alongside its unique ticker
                stock_volume_summary_index = ws.Cells(Rows.Count, 12).End(xlUp).Row + 1
                ws.Cells(stock_volume_summary_index, 12).Value = stock_volume
                
                ' --------------------
                '        BONUS
                ' --------------------
                ' Keep track of the largest percentage increase
                If IsNull(greatest_percentage_increase) Then
                    greatest_percentage_increase = yr_percent_change
                    greatest_percentage_increase_ticker = ws.Cells(i, 1)
                ElseIf greatest_percentage_increase < yr_percent_change Then
                    greatest_percentage_increase = yr_percent_change
                    greatest_percentage_increase_ticker = ws.Cells(i, 1)
                End If
                
                
                ' Kepp track of the largest percentage decrease
                If IsNull(greatest_percentage_decrease) Then
                    greatest_percentage_decrease = yr_percent_change
                    greatest_percentage_decrease_ticker = ws.Cells(i, 1)
                ElseIf greatest_percentage_decrease > yr_percent_change Then
                    greatest_percentage_decrease = yr_percent_change
                    greatest_percentage_decrease_ticker = ws.Cells(i, 1)
                End If
                
                
                ' Keep track of the largest volume
                If IsNull(greatest_total_vol) Then
                    greatest_total_vol = stock_volume
                    greatest_total_vol_ticker = ws.Cells(i, 1).Value
                ElseIf greatest_total_vol < stock_volume Then
                    greatest_total_vol = stock_volume
                    greatest_total_vol_ticker = ws.Cells(i, 1).Value
                End If
                        
                ' Updates | Resets
                yr_start_open = ws.Cells(i + 1, 3).Value
                stock_volume = 0
            End If
        Next i
        ' --------------------
        '        BONUS
        ' --------------------
        
        ws.Cells(2, 15).Value = greatest_percentage_increase_ticker
        ws.Cells(2, 16).Value = greatest_percentage_increase - 1
        
        ws.Cells(3, 15).Value = greatest_percentage_decrease_ticker
        ws.Cells(3, 16).Value = greatest_percentage_decrease - 1
        
        ws.Cells(4, 15).Value = greatest_total_vol_ticker
        ws.Cells(4, 16).Value = greatest_total_vol
        
    Next ws
End Sub

    


    

