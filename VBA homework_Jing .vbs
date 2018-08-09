Sub testing():

  Dim ws As Worksheet
  Dim total_stock_volume As Double
  total_stock_volume = 0
  Dim year_open As Double
  Dim year_close As Double
  Dim Greatest_total_volume As Double
  
  For Each ws In Worksheets
      lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
      Dim Summary_Table_Row As Double
      Summary_Table_Row = 2
      ws.Cells(1, 9).Value = "Ticker"
      ws.Cells(1, 12).Value = "Total Stock Volume"
      ws.Cells(1, 10).Value = "Yearly change"
      ws.Cells(1, 11).Value = "Percent change"
      ws.Cells(2, 14).Value = "Greatest % increase"
      ws.Cells(3, 14).Value = "Greatest % Decrease"
      ws.Cells(4, 14).Value = "Greatest total volume"
      ws.Cells(1, 15).Value = "Ticker"
      ws.Cells(1, 16).Value = "Value"
      ws.Range("K2:K" & lastrow).NumberFormat = "0.00%"
      ws.Cells(2, 16).NumberFormat = "0.00%"
      ws.Cells(3, 16).NumberFormat = "0.00%"
      'Initializing the first year_open value for the ticker in each sheet
      year_open = ws.Cells(2, 3).Value
    
    For i = 2 To lastrow
      If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
          ws.Range("I" & Summary_Table_Row).Value = ws.Cells(i, 1).Value
          total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
          ws.Range("L" & Summary_Table_Row).Value = total_stock_volume
          total_stock_volume = 0
          year_close = ws.Cells(i, 6).Value
          ws.Range("J" & Summary_Table_Row).Value = year_close - year_open
             If (year_close - year_open) > 0 Then
              ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
             Else
              ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
             End If
           'Checking if year_open is 0 or not, since Division by 0 is not possible
             If year_open <> 0 Then
                ws.Range("K" & Summary_Table_Row).Value = (year_close - year_open) / year_open
             Else
                ws.Range("K" & Summary_Table_Row).Value = 0
             End If
         Summary_Table_Row = Summary_Table_Row + 1
         year_open = ws.Cells(i + 1, 3).Value
      Else
         total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
      End If
    Next i

    ws.Cells(2, 16).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastrow))
    ws.Cells(3, 16).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastrow))
    ws.Cells(4, 16).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastrow))

    For i = 2 To lastrow
      If ws.Cells(i, 11).Value = ws.Cells(2, 16).Value Then
        ws.Cells(2, 15).Value = ws.Cells(i, 9).Value

      ElseIf ws.Cells(i, 11).Value = ws.Cells(3, 16).Value Then
        ws.Cells(3, 15).Value = ws.Cells(i, 9).Value
   
      ElseIf ws.Cells(i, 12).Value = ws.Cells(4, 16).Value Then
        ws.Cells(4, 15).Value = ws.Cells(i, 9).Value
      End If
   Next i

  Next ws

End Sub

