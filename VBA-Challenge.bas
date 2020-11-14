Attribute VB_Name = "Module1"
Sub stocks()
    For Each ws In Worksheets
    
        Dim row_count As Long
        Dim table_row_count As Long
        Dim total_stock_volume As Double
        Dim ticker As String
        Dim open_price As Double
        Dim close_price As Double
        Dim yearly_change As Double
        Dim percent_change As Double
        Dim last_row As Long
        Dim i As Long
        
        
        
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        table_row_count = 2
        
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % increase"
        ws.Cells(3, 15).Value = "Greatest % decrease"
        ws.Cells(4, 15).Value = "Greatest total volume"
        
        open_price = ws.Cells(2, 3).Value
        
        For i = 2 To last_row
            If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
                total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
            Else
                total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
                close_price = ws.Cells(i, 6).Value
                yearly_change = close_price - open_price
                If open_price > 0 Then
                    percent_change = yearly_change / open_price
                    ws.Cells(table_row_count, 11).Value = percent_change
                Else
                    ws.Cells(table_row_count, 11).Value = 0
                End If
        
                ticker = ws.Cells(i, 1).Value
                ws.Cells(table_row_count, 9).Value = ticker
                ws.Cells(table_row_count, 10).Value = yearly_change
                If yearly_change >= 0 Then
                    ws.Cells(table_row_count, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(table_row_count, 10).Interior.ColorIndex = 3
                End If
                
                ws.Cells(table_row_count, 11).NumberFormat = "0.00%"
                ws.Cells(table_row_count, 12).Value = total_stock_volume
                total_stock_volume = 0
                open_price = ws.Cells(i + 1, 3).Value
                table_row_count = table_row_count + 1
              End If
        Next i
        
        Dim min_percent As Double
        Dim max_percent As Double
        Dim temp_percent As Double
        Dim max_volume As Double
        Dim temp_volume As Double
        Dim min_max_count As Integer
        
        max_percent = -100
        
        For i = 2 To table_row_count
            temp_percent = ws.Cells(i, 11).Value
            If temp_percent > max_percent Then
                max_percent = temp_percent
                min_max_count = i
            End If
        Next i
        
        ws.Cells(2, 16).Value = ws.Cells(min_max_count, 9)
        ws.Cells(2, 17).Value = max_percent
        ws.Cells(2, 17).NumberFormat = "0.00%"
        
        min_percent = 100
        
        For i = 2 To table_row_count
            temp_percent = ws.Cells(i, 11).Value
            If temp_percent < min_percent Then
                min_percent = temp_percent
                min_max_count = i
            End If
        Next i
        
        ws.Cells(3, 16).Value = ws.Cells(min_max_count, 9)
        ws.Cells(3, 17).Value = min_percent
        ws.Cells(3, 17).NumberFormat = "0.00%"
        
        max_volume = 0
        
        For i = 2 To table_row_count
            temp_volume = ws.Cells(i, 12).Value
            If temp_volume > max_volume Then
                max_volume = temp_volume
                min_max_count = i
            End If
        Next i
        
        ws.Cells(4, 16).Value = ws.Cells(min_max_count, 9)
        ws.Cells(4, 17).Value = max_volume
        
     Next ws
              
End Sub


