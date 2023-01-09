Attribute VB_Name = "Module1"
Sub Summary()

'Prerequisites - data has been sorted by ticker and date(ascending)

'Define the variable, documenting the end point
Dim end_pint As Long
end_point = Range("A1").End(xlDown).Row

Range("I2") = "Ticker"
Range("J2") = "Yearly Change"
Range("K2") = "Percent Change"
Range("L2") = "Total Stock Volumn"

'Define the variable j. This is the end point of the result
Dim j As Integer
j = 3

Dim year_start As Double
year_start = 0
Dim year_end As Double
year_end = 0
Dim total_stock As Double
total_stock = 0

'Loop the data and summarise the yearly change and percenta change
For i = 2 To end_point
If Range("A" & i) <> Range("A" & (i - 1)) Then
Range("I" & j).Value = Range("A" & i)
year_start = Range("C" & i).Value
'Range("J" & j).Value = year_start
j = j + 1
ElseIf Range("A" & i) <> Range("A" & (i + 1)) Then
year_end = Range("F" & i).Value
'Range("K" & (j - 1)).Value = year_end
Range("J" & (j - 1)).Value = Format((year_end - year_start), "0.00")
If Range("J" & (j - 1)).Value > 0 Then
Range("J" & (j - 1)).Interior.ColorIndex = 4
Else
Range("J" & (j - 1)).Interior.ColorIndex = 3
End If
Range("K" & (j - 1)).Value = Format((year_end / year_start - 1), "0.00%")
End If
Next i

'Total Stock Volumn, using the sumif function
For t = 3 To (j - 1)
Range("L" & t) = WorksheetFunction.SumIf(Range("A2:A" & end_point), Range("I" & t), Range("G2:G" & end_point))
Next t

'Define the result table
Range("O3").Value = "Greatest % Increase"
Range("O4").Value = "Greatest % Decrease"
Range("O5").Value = "Greatest Total Volumn"
Range("P2").Value = "Ticker"
Range("Q2").Value = "Value"

'Define the initial values
Dim max_increase As Double
Dim max_decrease As Double
Dim max_vol As Double
max_increase = Range("K3").Value
max_decrease = Range("K3").Value
max_vol = Range("L3").Value

'Finding the max and min using for loop
For t = 3 To (j - 1)

If Range("K" & t).Value > max_increase Then
max_increase = Range("K" & t).Value
Range("P3").Value = Range("I" & t).Value
Range("Q3").Value = Format((Range("K" & t).Value), "0.00%")
End If

If Range("K" & t).Value < max_decrease Then
max_decrease = Range("K" & t).Value
Range("P4").Value = Range("I" & t).Value
Range("Q4").Value = Format((Range("K" & t).Value), "0.00%")
End If

If Range("L" & t).Value > max_vol Then
max_vol = Range("L" & t).Value
Range("P5").Value = Range("I" & t).Value
Range("Q5").Value = Range("L" & t).Value
End If

Next t

End Sub

