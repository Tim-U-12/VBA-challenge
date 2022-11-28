Attribute VB_Name = "Module1"
' Loops through all the stocks for one year and outputs the following information
' Ticker Symbol
' Yearly Change from the opening price at the beginning of a give year to the closing price at the end of the year
' Percentage change from the opening price at the beginning of a given year to the closing price at the end of the that year
' The total stock volume of the stock

Sub Collect_Ticker():
    ' loop index
    Dim i As Integer
    
    ' Final index
    Dim end_i As Integer
    end_i = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Titles for summary Table + inserting new column
    Range("i1").EntireColumn.Insert
    Range("i1").Value = "Ticker"
    
    Range("j1").EntireColumn.Insert
    Range("j1").Value = "Yearly Change"
    
    Range("k1").EntireColumn.Insert
    Range("k1").Value = "Percentage Change"
    
    Range("l1").EntireColumn.Insert
    Range("l1").Value = "Total Stock Volume"
    
    ' Iterate through A, finding all the unique tickers
    For i = 2 To end_i
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            ticker = Cells(Rows.Count, 9).End(xlUp).Row + 1
            Cells(ticker, 9).Value = Cells(i, 1).Value
        End If
    Next i
End Sub

