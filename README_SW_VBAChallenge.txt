VBA Code for Homework Due 6/27/2021
Stephanie Wilson

VBA Challenge Homework

Sub homeworktest()

'Declare variables
Dim Ticker As String
Dim Yearly_Change As Long
Dim Percent_Change As Long
Dim Total_Stock_Volume As Long
Dim open_price As Double
Dim close_price As Double
Dim volume_total As Long
Dim summary_row As Long


open_price = 0
close_price = 0
volume_total = 0
summary_row = 2


'Loop through each worksheet
For Each stock In Worksheets

'Add new column titles
stock.Cells(1, 9) = "Ticker"
stock.Cells(1, 10) = "Yearly_Change"
stock.Cells(1, 11) = "Percent_Change"
stock.Cells(1, 12) = "Total_Stock_Volume"

'Find lastrow in column
LastRow = stock.Cells(Rows.Count, 1).End(xlUp).Row

 ' Loop through rows in the column
  For i = 2 To LastRow

    ' Searches for when the value of the next cell is different than that of the current cell
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then

      'Add Ticker symbols
      Ticker = Cells(i, 1).Value
      
      'Add yearly change
      Yearly_Change = Cells(i - 1, 6).Value - Cells(i, 3).Value
      
      'Add percent change
      Percent_Change = Cells(i - 1, 6).Value - Cells(i, 3).Value / Cells(i, 3).Value

      ' Add to the volume Total
      volume_total = volume_total + Cells(i, 7).Value

      ' Print the summary total
      Range("I" & summary_row).Value = Ticker
      
      'Print the yearly change
      Range("J" & summary_row).Value = Yearly_Change
       If Range("J") >= 1 Then
        Range("J").Interior.Color = vbGreen
        Else: Range("J").Interior.Color = vbRed
      
      'Print the perecent change
      Range("K" & summary_row).Value = Percent_Change
      

      ' Print the volume amount to the total stock volume
      Range("L" & summary_row).Value = volume_total

      ' Add one to the summary table row
      summary_row = summary_row + 1
      
      ' Reset the volumne Total
      volume_total = 0

    Else

      ' Add to the Total Volume
      volume_total = volume_total + Cells(i, 7).Value


        End If
    
    Next i

Next stock
