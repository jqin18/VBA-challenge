Sub multi_year():
' loop through workbook
For Each ws In Worksheets

'set the position for the header
ws.Range("J1") = "Ticker"
ws.Range("K1") = "Yearly change"
ws.Range("L1") = "Percentage change"
ws.Range("M1") = "Total Volume"

' Set an initial variable for holding the ticker
  Dim ticker As String
  
  ' Set an initial variable for holding the total per ticker
  Dim Total_stock_vol As Double
  Total_stock_vol = 0

  ' Keep track of the location for each ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

' identify the last row
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'set the startvalue for the first ticker
startprice = ws.Cells(2, 3).Value

  ' Loop through all rows
  For i = 2 To LastRow

    ' Check if we are still within the same ticker, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the ticker
      ticker = ws.Cells(i, 1).Value
      
      'set the end price
      endprice = ws.Cells(i, 6).Value
      
      'set the yearly change
      yearly_change = endprice - startprice
      
      'set the percentage changed (elimnate the zero division error)
     If startprice <> 0 Then
      percentage_changed = FormatPercent((yearly_change / startprice))
      Else
      percentage_changed = 0
      
      End If

      ' Add to the ticker Total
      Total_stock_vol = Total_stock_vol + ws.Cells(i, 7).Value

      ' Print the ticker in the Summary Table
      ws.Range("J" & Summary_Table_Row).Value = ticker

      ' Print the ticker total Amount to the Summary Table
      ws.Range("M" & Summary_Table_Row).Value = Total_stock_vol
      
      'Print yearly_change
      ws.Range("K" & Summary_Table_Row).Value = yearly_change
      
      'Print percentage changed
      ws.Range("L" & Summary_Table_Row).Value = percentage_changed

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the stock volume Total
      Total_stock_vol = 0
      
      'reset the startprice
      startprice = ws.Cells(i + 1, 3)

    ' If the cell immediately following a row is the same ticker...
    Else

      ' Add to the ticker Total
      Total_stock_vol = Total_stock_vol + ws.Cells(i, 7).Value

    End If
        
    'formatting based on the yearly change
    If ws.Cells(i, 11).Value > 0 Then
        ws.Cells(i, 11).Interior.ColorIndex = 4

   ElseIf ws.Cells(i, 11).Value < 0 Then
        ws.Cells(i, 11).Interior.ColorIndex = 3
        
    End If
    

  Next i
  Next ws

End Sub
