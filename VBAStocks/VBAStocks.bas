Attribute VB_Name = "Module1"

Sub stockdata():

Dim ws As Worksheet

For Each ws In Worksheets

    Dim open_price As Double
    Dim close_price As Double
    Dim last_row As Long
    Dim last_column As Long
    Dim ticker As String
    Dim percentchange As Double
    Dim yearlychange As Double
    Dim total_stockvolume As Double
    Dim max_ticker As String
    Dim min_ticker As String
    Dim max_percent As Double
    Dim min_percent As Double
    Dim max_volume_ticker As String
    Dim max_volme As Double
    
    
    Dim Summary_Table_Row As Long
    Summary_Table_Row = 2
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    total_stockvolume = 0
    ticker = " "
    yearlychange = 0
    percentchange = 0
    close_price = 0
    open_price = ws.Cells(2, 3).Value
    max_ticker = " "
    min_ticker = " "
    max_percent = 0
    min_percent = 0
    max_volume_ticker = " "
    max_volume = 0
    
    ws.Cells(1, 10).Value = "Ticker"
    ws.Cells(1, 11).Value = "Yearly Change"
    ws.Cells(1, 12).Value = "Percent Change"
    ws.Cells(1, 13).Value = "Total Stock Volume"
    ws.Cells(2, 16).Value = "Greatest % Increase"
    ws.Cells(3, 16).Value = "Greatest % Decrease"
    ws.Cells(4, 16).Value = "Greatest Total Volume"
    ws.Cells(1, 17).Value = "Ticker"
    ws.Cells(1, 18).Value = "Value"

For i = 2 To last_row
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    ticker = ws.Cells(i, 1).Value
    close_price = ws.Cells(i, 6).Value
    yearlychange = close_price - open_price
        If open_price <> 0 Then
            percentchange = (yearlychange / open_price)
            
        Else
            percentchange = 0
            yearlychange = 0
        End If
        
        
    ws.Range("J" & Summary_Table_Row).Value = ticker
    ws.Range("K" & Summary_Table_Row).Value = yearlychange
    ws.Range("L" & Summary_Table_Row).Value = percentchange
    ws.Range("M" & Summary_Table_Row).Value = total_stockvolume
    ws.Columns("L").NumberFormat = "0.00%"
    
        If ws.Range("K" & Summary_Table_Row).Value > 0 Then
            ws.Cells(Summary_Table_Row, 11).Interior.ColorIndex = 4
            
        ElseIf ws.Range("K" & Summary_Table_Row).Value <= 0 Then
            ws.Cells(Summary_Table_Row, 11).Interior.ColorIndex = 3
        End If
        
        If (percentchange > max_percent) Then
            max_percent = percentchange
            max_ticker = ticker
        ElseIf (percentchange < min_percent) Then
            min_percent = percentchange
            min_ticker = ticker
        End If
        If (total_stockvolume > max_volume) Then
            max_volume = total_stockvolume
            max_volume_ticker = ticker
        End If
        
    ws.Cells(2, 17).Value = max_ticker
    ws.Cells(3, 17).Value = min_ticker
    ws.Cells(4, 17).Value = max_volume_ticker
    ws.Cells(2, 18).Value = max_percent
    ws.Cells(3, 18).Value = min_percent
    ws.Cells(4, 18).Value = max_volume
    ws.Cells(2, 18).NumberFormat = "0.00%"
    ws.Cells(3, 18).NumberFormat = "0.00%"
    
    Summary_Table_Row = Summary_Table_Row + 1
    
    
    total_stockvolume = 0
    close_price = 0
    yearlychange = 0
    percentchange = 0
    open_price = ws.Cells(i + 1, 3).Value
    
    Else
    
    total_stockvolume = total_stockvolume + ws.Cells(i, 7).Value

    End If
    
    
Next i


Next ws




End Sub
