Attribute VB_Name = "Module1"
Sub stock_market()

Dim ticker
Dim open_price
Dim close_price
Dim total_stock_volume As Variant
Dim yearly_change
Dim percent_change
Dim Summary_Table_Row
Dim wsCount As Integer


Dim greatest_increase_ticker
Dim greatest_decrease_ticker
Dim greatest_volume_ticker
Dim greatest_increase_value
Dim greatest_decrease_value
Dim greatest_volume_value

greatest_increase_ticker = "test"
greatest_decrease_ticker = "test"
greatest_volume_ticker = "test"
greatest_increase_value = 0
greatest_decrease_value = 0
greatest_volume_value = 0


wsCount = ActiveWorkbook.Worksheets.Count

For ws = 1 To wsCount
    Worksheets(ws).Range("I1") = "Ticker"
    Worksheets(ws).Range("J1") = "Yearly Change"
    Worksheets(ws).Range("K1") = "Percent Change"
    Worksheets(ws).Range("L1") = "Total Stock Volume"
    
    Worksheets(ws).Range("P1") = "Ticker"
    Worksheets(ws).Range("Q1") = "Value"
    Worksheets(ws).Range("O2") = "Greatest % Increase"
    Worksheets(ws).Range("O3") = "Greatest % Decrease"
    Worksheets(ws).Range("O4") = "Greatest Total Volume"
    Worksheets(ws).Range("Q2:Q3").NumberFormat = "0.00%"
    
    Summary_Table_Row = 2
    open_price = 2
    
    For I = 2 To Worksheets(ws).Cells(Rows.Count, 1).End(xlUp).Row
        If Worksheets(ws).Cells(I, 1) <> Worksheets(ws).Cells((I + 1), 1) Then
        
        ticker = Worksheets(ws).Cells(I, 1)
        
        total_stock_volume = total_stock_volume + Cells(I, 7)
        Worksheets(ws).Range("I" & Summary_Table_Row).Value = ticker
        
        close_price = Worksheets(ws).Cells(I, 6)
       
        yearly_change = close_price - Worksheets(ws).Cells(open_price, 3)
        
        Worksheets(ws).Range("J" & Summary_Table_Row).Value = yearly_change
        
        
        
        
        If yearly_change > 0 Then
        Worksheets(ws).Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
        Else
        Worksheets(ws).Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
        End If
        
        If Worksheets(ws).Cells(open_price, 3) <> 0 Then
            
            percent_change = yearly_change / Worksheets(ws).Cells(open_price, 3)
            
        Else
            percent_change = 0
        End If
        
         Worksheets(ws).Range("K" & Summary_Table_Row).Value = percent_change
         Worksheets(ws).Range("K" & Summary_Table_Row).NumberFormat = "00.0%"
         
         Worksheets(ws).Range("L" & Summary_Table_Row).Value = total_stock_volume
         
         If total_stock_volume > greatest_volume_value Then
            greatest_volume_value = total_stock_volume
            greatest_volume_ticker = ticker
         End If
         
         Summary_Table_Row = Summary_Table_Row + 1
         open_price = I + 1
         total_stock_volume = 0
         '-----------------------
         Else
         total_stock_volume = total_stock_volume + Cells(I, 7)
            
         
         End If
         
         If percent_change > greatest_increase_value Then
         greatest_increase_value = percent_change
         greatest_increase_ticker = ticker
         ElseIf percent_change < greatest_decrease_value Then
         greatest_decrease_value = percent_change
         greatest_decrease_ticker = ticker
         End If
         
        
    
    
    Next I
    
    Worksheets(ws).Range("P2") = greatest_increase_ticker
    Worksheets(ws).Range("Q2") = greatest_increase_value
    Worksheets(ws).Range("P3") = greatest_decrease_ticker
    Worksheets(ws).Range("Q3") = greatest_decrease_value
    Worksheets(ws).Range("P4") = greatest_volume_ticker
    Worksheets(ws).Range("Q4") = greatest_volume_value

   
    greatest_increase_value = 0
    greatest_decrease_value = 0
    greatest_volume_value = 0

    







Next ws

End Sub



