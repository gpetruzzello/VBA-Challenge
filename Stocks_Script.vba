Sub wallstreet()

' Loop through all sheets

  For Each ws In Worksheets
  
  ' Name summary headers
  
  ws.Range("J1").Value = "Ticker"
  ws.Range("K1").Value = "Yearly Change"
  ws.Range("L1").Value = "Percent Change"
  ws.Range("M1").Value = "Total Stock Volume"
  
  ' Set an initial variable for holding the ticker
  Dim Ticker As String
  ' Set an initial variable for holding the total per ticker
  Dim Total_Stock_Volume As Double
 Total_Stock_Volume = 0
  ' Keep track of the location for each ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
  
  ' Loop through all rows
  
        For I = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
  
    ' Check if we are still within the same ticker, if it is not...
    
            Ticker = ws.Cells(I, 1).Value
    
             next_row_ticker = ws.Cells(I + 1, 1).Value
    
            Closing_Price = ws.Cells(I, 6).Value
    'Opening_Price = Cells(i, 3).Value
    
    
    
            previous_row_ticker = ws.Cells(I - 1, 1).Value
    
    
            If previous_row_ticker <> Ticker Then
                Opening_price = ws.Cells(I, 3).Value

    
            ElseIf next_row_ticker = Ticker Then
      ' Add to the Total_Stock_Volume
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(I, 7).Value

            Else
      ' Add to the Total_Stock_Volume
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(I, 7).Value
      ' Print the Ticker in the Summary Table
                ws.Range("J" & Summary_Table_Row).Value = Ticker
      ' Print the Total Stock Volume to the Summary Table
                ws.Range("M" & Summary_Table_Row).Value = Total_Stock_Volume
      
      
      
                Yearly_Change = Closing_Price - Opening_price
      
                ws.Range("K" & Summary_Table_Row).Value = Yearly_Change
                
                If Yearly_Change <> 0 Then
                
                    ws.Range("L" & Summary_Table_Row).Value = FormatPercent(Yearly_Change / Opening_price)
                 End If
                
                 
                  If ws.Range("L" & Summary_Table_Row).Value < 0 Then
                    ws.Range("L" & Summary_Table_Row).Interior.ColorIndex = 3
                   End If
                 
                 If ws.Range("L" & Summary_Table_Row).Value > 0 Then
                    ws.Range("L" & Summary_Table_Row).Interior.ColorIndex = 4
                   End If
                 

      ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Total_Stock_Volume
                Total_Stock_Volume = 0

    
          End If
 
         Next I
    
    Next

End Sub