Attribute VB_Name = "Module1"
Sub stocks()

For Each ws In Worksheets

    Dim ticker As String
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim stock_volume As Double
    Dim open_total As Double
    Dim close_total As Double
    Dim ticker_row As Integer
    
    ws.Cells(1, 9).Value = "ticker"
    ws.Cells(1, 10).Value = "yearly change"
    ws.Cells(1, 11).Value = "percent change"
    ws.Cells(1, 12).Value = "total stock volume"
    
  ticker_row = 2
       open_total = ws.Cells(2, 3).Value
       For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                stock_volume = stock_volume + ws.Cells(i, 7).Value
                ws.Cells(ticker_row, 12).Value = stock_volume
                ws.Cells(ticker_row, 9).Value = ticker
            
                stock_volume = 0
            
                close_total = ws.Cells(i, 6).Value
                yearly_change = close_total - open_total
                ws.Cells(ticker_row, 10).Value = yearly_change
        
                
                If open_total = 0 Then
                percent_change = (close_total - open_total)
                Else
                percent_change = (close_total - open_total) / (open_total)
                End If
                
                ws.Cells(ticker_row, 11).Value = percent_change
                ws.Cells(ticker_row, 11).NumberFormat = "0.00%"
                
                open_total = ws.Cells(i, 3).Value
                yearly_change = 0
                ticker_row = ticker_row + 1
                
                
        Else
            stock_volume = stock_volume + ws.Cells(i, 7).Value
          
        End If
        
         If ws.Cells(i, 10).Value < 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 3
         End If
        
        If ws.Cells(i, 10).Value > 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 4
        
        End If
        
        Next i
        
        
    Next ws

End Sub



