Attribute VB_Name = "Module1"
Sub Stock()

  ' Set an initial variable for holding the Stock Ticker Symbol
  Dim Stock_Ticker As String

  ' Set an initial variable for holding the total stock volume
  Dim Volume_Total As Double
  Volume_Total = 0
  
  Dim lastrow As Double
  

  ' Keep track of the location for each Stock Ticker in the summary table
  Dim Summary_Table_Row As Double
  Summary_Table_Row = 2

    lastrow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "G").End(xlUp).Row
    
  ' Loop through all Stock Ticker Volume
  For i = 2 To lastrow

    ' Check if we are still within the same stock ticker, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Stock Ticker name
      Stock_Ticker = Cells(i, 1).Value

      ' Add to the Stock Volume Total
      Volume_Total = Volume_Total + Cells(i, 7).Value

      ' Print the Stock Volume Total in the Summary Table
      Range("I" & Summary_Table_Row).Value = Stock_Ticker

      ' Print the Volume Amount to the Summary Table
      Range("J" & Summary_Table_Row).Value = Volume_Total

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Stock Volume Total
      Volume_Total = 0

    ' If the cell immediately following a row is the same stock ticker symbol...
    Else

      ' Add to the Volume Total
      Volume_Total = Volume_Total + Cells(i, 7).Value

    End If

  Next i

End Sub



