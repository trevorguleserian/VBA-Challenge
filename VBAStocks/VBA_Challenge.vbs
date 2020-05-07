Sub vba_challenge()

for each ws in Worksheets

  
  Dim Ticker_Symbol As String
  Dim Change_Price as Double
  Dim Percent_Change as Double
  Dim Open_Price as Double
  Dim Close_Price as Double
  Dim Stock_Volume As Double
  Dim NumRows as Double
  Dim NumRows2 as Double
  Dim Summary_Table_Row As Integer
  Stock_Volume = 0
  Summary_Table_Row = 2
  NumRows = Range("A2", Range("A2").End(xldown)).Rows.Count
  

  For i = 2 to NumRows 
    
  
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      
      Ticker_Symbol = ws.Cells(i, 1).Value

      
      Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
      
      Close_Price = ws.cells(i,6).value 
      Open_Price = ws.cells(i,3).value
    
      Change_Price = Close_Price - Open_Price 
      
      Percent_Change = (Change_Price/Open_Price)*100

      ws.cells(1,9).value = "Ticker Symbol"
      ws.cells(1,10).value = "Change in Price"
      ws.cells(1,11).value = "Percent Change"
      ws.cells(1,12).value = "Stock Volume"
    




      ws.Range("I" & Summary_Table_Row).Value = Ticker_Symbol

      ws.Range("J" & Summary_Table_Row).Value = Change_Price

      ws.Range("K" & Summary_Table_Row).Value = Percent_Change

     ws. Range("L" & Summary_Table_Row).Value = Stock_Volume

        if  ws.Range("J" & Summary_Table_Row).Value >= 0 Then

            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
        Else 

            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3

        End If

        if ws.Range("K" & Summary_Table_Row).Value >= 0 Then

           ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
        Else 

            ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3

        End If

      Summary_Table_Row = Summary_Table_Row + 1
      
      Stock_Volume = 0

   
    Else

      Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
    
    End If

  Next i

Next ws

End Sub
