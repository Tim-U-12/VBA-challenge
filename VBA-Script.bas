Attribute VB_Name = "Module1"
' Loops through all the stocks for one year and outputs the following information:
' Percentage change from the opening price at the beginning of a given year to the closing price at the end of the that year
' Done - Ticker Symbol
' Done - Yearly Change from the opening price at the beginning of a give year to the closing price at the end of the year
' Done - The total stock volume of the stock

Sub Collect_Ticker():
    ' loop index
    Dim i As Integer
    
    ' end index
    Dim end_i As Integer
    end_i = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' init year start open
    Dim yr_start_open As Double
    yr_start_open = Cells(2, 3).Value
    
    ' Yearly Change
    Dim yr_change As Double
    
    ' Yearly Percentage Change
    Dim yr_percent_change As Double
    
    ' #############################################################################################
    
    ' Titles for summary Table + inserting new column
    Range("i1").EntireColumn.Insert
    Range("i1").Value = "Ticker"
    
    Range("j1").EntireColumn.Insert
    Range("j1").Value = "Yearly Change"
    
    Range("k1").EntireColumn.Insert
    Range("k1").Value = "Percentage Change"
    
    Range("l1").EntireColumn.Insert
    Range("l1").Value = "Total Stock Volume"
    
    ' Change width of summary columns
    Columns("J:M").ColumnWidth = 20
    
    ' Change the format of column k to be in percentages
    Columns("k").NumberFormat = "0.00%"
    
    ' #############################################################################################
    
    stock_volume = 0
    ' Iterate through the rows in a single worksheet
    For i = 2 To end_i
        stock_volume = stock_volume + Cells(i, 7).Value
    
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            ' Adding unique ticker to summary table
            ticker = Cells(Rows.Count, 9).End(xlUp).Row + 1
            Cells(ticker, 9).Value = Cells(i, 1).Value
            
            ' Adding the difference between open and close to the summary table
            yr_change = Cells(i, 6).Value - yr_start_open
            yr_change_index = Cells(Rows.Count, 10).End(xlUp).Row + 1
            Cells(yr_change_index, 10).Value = yr_change
            
            ' Format the yearly change outcomes
            If yr_change < 0 Then
                Cells(yr_change_index, 10).Interior.ColorIndex = 3
            Else
                Cells(yr_change_index, 10).Interior.ColorIndex = 4
            End If
            
            ' Percentage change from the opening price at the beginning of a given year to end
            yr_percent_change = Cells(i, 6).Value / yr_start_open
            yr_percent_change_index = Cells(Rows.Count, 11).End(xlUp).Row + 1
            Cells(yr_percent_change_index, 11).Value = yr_percent_change - 1
                            
            ' Adding the volume stock to the summary table alongside its unique ticker
            stock_volume_summary_index = Cells(Rows.Count, 12).End(xlUp).Row + 1
            Cells(stock_volume_summary_index, 12).Value = stock_volume
            
            yr_start_open = Cells(i + 1, 3).Value
            stock_volume = 0
        End If
    Next i
End Sub

