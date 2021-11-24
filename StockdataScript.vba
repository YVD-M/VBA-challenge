Sub Stockdata()

  Dim Ticker_Name As String
  Dim Yearly_Change As Double
  Dim Percent_Change As Double
  Dim Open_Price As Double
  Dim Close_Price As Double
  
  Range("I1").Value = "Ticker"
  Range("J1").Value = "Yearly Change"
  Range("K1").Value = "Percent Change"
  Range("L1").Value = "Total Stock Volume"


  Total_Stock_Volume = 0

  Dim Summary_Section_Row As Integer
  
  Summary_Section_Row = 2
  
  LastRow = Cells(Rows.Count, 1).End(xlUp).Row
  
  
  Open_Price = Cells(2, 3).Value
  

  For i = 2 To LastRow


    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then


      Ticker_Name = Cells(i, 1).Value

      
      Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value

    
      Range("I" & Summary_Section_Row).Value = Ticker_Name


      Range("L" & Summary_Section_Row).Value = Total_Stock_Volume

      
      Close_Price = Cells(i, 6).Value
      Yearly_Change = Close_Price - Open_Price
      Cells(Summary_Section_Row, 10).Value = Yearly_Change
      
      
      If Open_Price <> 0 Then
      Percent_Change = (Yearly_Change / Open_Price)
      Cells(Summary_Section_Row, 11).Value = Round((Percent_Change * 100), 2) & "%"
      
      End If
      
      
      Summary_Section_Row = Summary_Section_Row + 1
      Total_Stock_Volume = 0
      Open_Price = Cells(i + 1, 3).Value
      
      
    
    Else

      Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
      
      End If
      
      
    
        

  Next i

End Sub


