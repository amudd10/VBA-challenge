Sub Stock_market()

    'Worked with Georgia Myers
    'Defining everything
    Dim ws As Worksheet
    For Each ws In Worksheets
        ws.Activate
    Dim ticker As String
    Dim vol As LongLong
    Dim Summary_Table_Row As Integer
    Dim start_price As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    start_price = Cells(2, 3).Value
    'Setup integers for loop
    Summary_Table_Row = 2
    
    'For loop to loop through every worksheet'
    'For Each ws In Worksheets
    
        'vol = 0
        'row_start = 2
        'yearly_percentage = 0
        'total_stock_volume = 0
    
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
    
    
    
    'Loop
        For i = 2 To Cells(Rows.Count, 2).End(xlUp).Row
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            
                'Find all values
                ticker = Cells(i, 1).Value
                vol = vol + Cells(i, 7).Value
                yearly_change = Cells(i, 6).Value - start_price
                percent_change = (yearly_change / start_price)
                
                'Insert value into summary
                Range("I" & Summary_Table_Row).Value = ticker
                Range("J" & Summary_Table_Row).Value = yearly_change
                    If yearly_change >= 0 Then
                     Range("J" & Summary_Table_Row).Interior.Color = RGB(0, 128, 0)
                    Else
                     Range("J" & Summary_Table_Row).Interior.Color = RGB(255, 0, 0)
                    End If
                Range("K" & Summary_Table_Row).Value = percent_change
                Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                Range("L" & Summary_Table_Row).Value = vol
                vol = 0
                Summary_Table_Row = Summary_Table_Row + 1
                start_price = Cells(i + 1, 3).Value
            Else
                vol = vol + Cells(i, 7).Value
            End If
            
        Next i
        For i = 2 To Cells(Rows.Count, 10).End(xlUp).Row
            Cells(2, 15) = "Greatest % Increase"
            Cells(3, 15) = "Greatest % Decrease"
            Cells(4, 15) = "Greatest Total Volume"
            Cells(1, 16) = "Ticker"
            Cells(1, 17) = "Value"
            Cells(2, 17).NumberFormat = "0.00%"
            Cells(3, 17).NumberFormat = "0.00%"
            Cells(2, 17).Value = WorksheetFunction.Max(Range("k2:k" & Cells(Rows.Count, 10).End(xlUp).Row))
            
            If Cells(i, 11).Value = Cells(2, 17).Value Then
                Cells(2, 16).Value = Cells(i, 9).Value
            End If
            
            Cells(3, 17).Value = WorksheetFunction.Min(Range("k2:k" & Cells(Rows.Count, 10).End(xlUp).Row))
            If Cells(i, 11).Value = Cells(3, 17).Value Then
                Cells(3, 16).Value = Cells(i, 9).Value
            End If
            
            Cells(4, 17).Value = WorksheetFunction.Max(Range("l2:l" & Cells(Rows.Count, 10).End(xlUp).Row))
            If Cells(i, 12).Value = Cells(4, 17).Value Then
                Cells(4, 16).Value = Cells(i, 9).Value
            End If
            
        Next i
        Next ws
            
        
      
             
End Sub


