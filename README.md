# VBA-Challenge
'Set an initial varible for holding the ticker
  Dim Ticker As String
  'Set an initial varible for holding yearly change
  Dim Yearly_Change As CalculatedItems
  
  'Set an initial varible for holding percent change
  Dim Percent_Change As CalculatedItems
  
  'Set an initial varible for holding the the total per ticker
  Dim Total_Stock_volume As Double
  Total_Stock_volume = 0
  
  'Keep track of the location for each ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 4

  'Loop Through all ticker purchases
  For i = 2 To 22771
  
  'Check if we are still within the same ticker, if it is not...
  If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
  
  'Set the Ticker name
  Ticker_Name = Cells(i, 1).Value
  
  'Add to the Yearly Change
  Yearly_Change = Yearly_Change + Cells(i, 2).Value
  
  'Add to Percent Change
  Percent_Change = Percent_Change + Cells(i, 3).Value
  
  'Add to Total Stock Volume
  Total_Stock_volume = Total_Stock_volume + Cells(i, 7).Value
  
  'Print the Ticker in the Summary Table
  Range("J" & Summary_Table_Row).Value = Ticker_Name
  
  'Print the Yearly Change in the Summary Table
  Range("K" & Summary_Table_Row).Value = Yearly_Change
  
  'Print the percent change in the Summary Table
  Range("L" & Summary_Table_Row).Value = Percent_Change
  
  'Print the Total Stock Volume to the Summary Table
  Range("M" & Summary_Table_Row).Value = Total_Stock_volume
  
  'Reset the Yearly Change
  Yearly_Change = 0
  
  'If the cell immediately following a row is the same ticker...
  Else
  
  'Add to the Yearly Change
 Yearly_Change = Yearly_Change + Cells(i, 2).Value
  
  'Reset Percent Change
  Percent_Change = 0
  

  'Add to the Percent Change
  Percent_Change = Percent_Change + Cells(i, 3).Value
  
  
  'Reset the Total Stock Volume
  Total_Stock_volume = 0

  
  'Add to the Total Stock Volume
  Total_Stock_volume = Total_Stock_volume + Cells(i, 7).Value
  
  End If
  
 Next i
  
End Sub
