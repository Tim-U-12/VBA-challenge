Attribute VB_Name = "Module1"
' Loops through all the stocks for one year and outputs the following information:
' Yearly Change from the opening price at the beginning of a give year to the closing price at the end of the year
' Percentage change from the opening price at the beginning of a given year to the closing price at the end of the that year
' Done - Ticker Symbol
' Done - The total stock volume of the stock

Sub Collect_Ticker():
    ' loop index
    Dim i As Integer
    
    ' end index
    Dim end_i As Integer
    end_i = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Yearly stock open price
    Dim yr_stock_open As Double
    
    ' Yearly stock close price
    Dim yr_stock_close As Double
    
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
    
    
    stock_volume = 0
    
    ' Iterate through the rows in a single worksheet
    For i = 2 To end_i
        stock_volume = stock_volume + Cells(i, 7).Value
    
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            ' Adding unique ticker to summary table
            ticker = Cells(Rows.Count, 9).End(xlUp).Row + 1
            Cells(ticker, 9).Value = Cells(i, 1).Value
            
            ' Adding the volume stock to the summary table alongside its unique ticker
            stock_volume_summary_index = Cells(Rows.Count, 12).End(xlUp).Row + 1
            Cells(stock_volume_summary_index, 12).Value = stock_volume
            stock_volume = 0
        End If
    Next i
End Sub

