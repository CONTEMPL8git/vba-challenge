Attribute VB_Name = "Module1"
Sub stocks():
Dim ws As Worksheet
For Each ws In Worksheets
ws.Activate

Dim ticker As String
Dim total_stock_volume As LongLong
total_stock_volume = 0
'Keep track of the location for each ticker in the summary table
Dim Summary_Row As Integer
Summary_Row = 2

Dim open_price As Double
'MsgBox (Cells(2, 3).Value)

LastRow = Cells(Rows.Count, 1).End(xlUp).Row

Dim close_price As Double
Dim yearly_change As Double
Dim percentage_change As Double

Range("I1").Value = "Ticker"
 Range("J1").Value = "Yearly Change"
 Range("K1").Value = "Percent Change"
 Range("L1").Value = "Total Stock Volume"
  Range("O2").Value = "Greatest % Increase"
  Range("O3").Value = "Greatest % Decrease"
  Range("O4").Value = "Greatest Total Volume"
  Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
'Loop through all tickers
For i = 2 To LastRow - 1
 
    ' Check if we are still within the same ticker, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

'open price set***
    open_price = Range("C" & Summary_Row).Value
      
      ' Set the ticker name
      ticker = Cells(i, 1).Value
      
      ' Add to the total_stock_volume
      total_stock_volume = total_stock_volume + Cells(i, 7).Value
      
      ' Print the ticker in the Summary Table
      Range("I" & Summary_Row).Value = ticker
      
      ' Print the total_stock_volume to the Summary Table
      Range("L" & Summary_Row).Value = total_stock_volume
      
      ' do the yearly change calculation, close price minus open price
      close_price = Cells(i, 6).Value
      yearly_change = close_price - open_price
      Range("J" & Summary_Row).Value = yearly_change
      
      'do the percentage change, its yearly change divided by open price
      'need to multiply times 100 also
        percentage_change = (yearly_change / open_price)
        Range("K" & Summary_Row).Value = percentage_change
        Range("K" & Summary_Row).NumberFormat = "0.00%"
        
      ' Add one to the summary row
      Summary_Row = Summary_Row + 1
      
      ' Reset the total_stock_volume
      total_stock_volume = 0
      
      ' If the cell immediately following a row is the same ticker...
    Else
      ' Add to the total_stock_volume
      total_stock_volume = total_stock_volume + Cells(i, 7).Value
    End If
    Next i
    
'conditional coloring of cells if neg red, pos green
'***NEED NEW LETTER CAN'T USE i again
For e = 2 To LastRow - 1
    If Cells(e, 10).Value < 0 Then
        Cells(e, 10).Interior.ColorIndex = 3
    ElseIf Cells(e, 10).Value > 0 Then
        Cells(e, 10).Interior.ColorIndex = 4
    End If
  Next e
  
For g = 2 To LastRow - 1
    If Cells(g, 11).Value < 0 Then
        Cells(g, 11).Interior.ColorIndex = 3
    ElseIf Cells(g, 11).Value > 0 Then
        Cells(g, 11).Interior.ColorIndex = 4
    End If
Next g
  
'Greatest increase code
For y = 2 To Summary_Row - 2
    If Cells(y + 1, 11).Value > Cells(y, 11).Value Then
    Cells(2, 17).Value = Cells(y + 1, 11).Value
    Cells(2, 16).Value = Cells(y + 1, 9).Value
    Range("Q2").NumberFormat = "0.00%"
    
End If

    Next y
    
'Greatest decrease code
For v = 2 To Summary_Row - 2
    If Cells(v + 1, 11).Value < Cells(v, 11).Value Then
    Cells(3, 17).Value = Cells(v + 1, 11).Value
    Cells(3, 16).Value = Cells(v + 1, 9).Value
    Range("Q3").NumberFormat = "0.00%"
    
End If
    Next v

'Greatest Total Volume
For c = 2 To Summary_Row - 2
    If Cells(c + 1, 12).Value > Cells(c, 12).Value Then
    Cells(4, 17).Value = Cells(c + 1, 12).Value
    Cells(4, 16).Value = Cells(c + 1, 9).Value

    
End If
    Next c

Next ws

End Sub












































