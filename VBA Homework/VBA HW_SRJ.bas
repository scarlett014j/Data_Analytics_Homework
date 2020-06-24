Attribute VB_Name = "Module1"
Sub Dosomething()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call test4
    Next
    Application.ScreenUpdating = True
End Sub
Sub test4()

'checks for the last row of data
Dim last_row As Long
    last_row = Cells(Rows.count, 1).End(xlUp).Row

Dim TSV As Long
Dim close_price As Double
Dim open_price As Double
Dim year_change As Double
Dim percent_change As Double
Dim ticker As String
Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
Dim count As Integer

Range("k:k").NumberFormat = "0.00%"
Range("P2", "P3").NumberFormat = "0.00%"

  
Range("P1").Value = "Value"
Range("O1").Value = "Ticker"
Range("N2").Value = "Greatest % Increase"
Range("N3").Value = "Greatest % Decrease"
Range("N4").Value = "Greatest Total Volume"
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

TSV = 0
close_price = 0
open_price = 0
year_change = 0
percent_change = 0
count = 0


For r = 2 To last_row
    'when the ticker symbols are different
    If Cells(r + 1, 1).Value <> Cells(r, 1).Value Then
        ticker = Cells(r, 1).Value
        TSV = Cells(r, 7).Value + TSV
        open_price = Cells(r - count, 3).Value
        close_price = Cells(r, 6).Value
        year_change = close_price - open_price
        'to account for when the opening price is zero (I used the closing price for the same day since it was not clear what to do for that particular instance)
       If open_price > 0 Then
            percent_change = (year_change / open_price)
       Else
           percent_change = (year_change)
           
       End If
       'Header formatting
        Range("I" & Summary_Table_Row).Value = ticker
        Range("J" & Summary_Table_Row).Value = year_change
        Range("K" & Summary_Table_Row).Value = percent_change
        Range("L" & Summary_Table_Row).Value = TSV
        
        Summary_Table_Row = Summary_Table_Row + 1
        
        'Resetting the variables for the next loop
        TSV = 0
        close_price = 0
        open_price = 0
        year_change = 0
        percent_change = 0
        count = 0
    Else
        'calculating the total stock volume
        TSV = TSV + Cells(r, 7).Value
        count = count + 1
   End If
   
   TSV = 0
Next r
   
'This block is the color formatting
For I = 2 To Summary_Table_Row - 1
    If Cells(I, 10) > 0 Then
        Cells(I, 10).Interior.ColorIndex = 10
    Else
        Cells(I, 10).Interior.ColorIndex = 3
        
    End If
  Next I

'this block finds the greatest % increase and decrease and the largest total stock volume
Dim max As Double
Dim min As Double
Dim biggest As Long



max = 0
min = 0
biggest = 0

For j = 2 To Summary_Table_Row
    If Cells(j, 11).Value > max Then
        max = Cells(j, 11).Value
        Range("P2") = max
        Cells(2, 15).Value = Cells(j, 9).Value
        
        
      
    End If
Next j

For k = 2 To Summary_Table_Row
    If Cells(k, 11).Value < min Then
        min = Cells(k, 11).Value
        Range("P3") = min
        Cells(3, 15).Value = Cells(k, 9).Value
        
    End If

Next k

 For a = 2 To Summary_Table_Row
    If Cells(a, 12).Value > biggest Then
        biggest = Cells(a, 12).Value
        Range("P4") = biggest
        Cells(4, 15).Value = Cells(a, 9).Value
      
    End If
 
Next a

              
        
        

End Sub




