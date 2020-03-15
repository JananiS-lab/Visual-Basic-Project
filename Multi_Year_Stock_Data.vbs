Attribute VB_Name = "Module1"
Sub stockdata():
Dim ticker_name As String
Dim ticker_value As String
Dim ticker_total As Double
Dim Summary_table_row As Integer
Dim x As Integer
Dim Max_1 As Double
Dim Min_1 As Double
Dim Growth As Double
Dim Percent_Growth As Double
Dim Cond1 As FormatCondition, Cond2 As FormatCondition
Dim rg As Range
Set rg = Range("J2", Range("J2").End(xlDown))
Dim Max As Double
Dim Min As Double
Dim Max_Total As Double
Summary_table_row = 2
Dim LastRow As Long
Set sht = ActiveSheet
LastRow = sht.Cells(sht.Rows.Count, "A").End(xlUp).Row - 1
' This is a loop to run from first to last row, provided the condition is satisfied
For I = 2 To LastRow
    If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
        ticker_name = Cells(I, 1).Value
        ticker_total = ticker_total + Cells(I, 7).Value
        Range("I" & Summary_table_row).Value = ticker_name
        Range("L" & Summary_table_row).Value = ticker_total
        Min_1 = Cells((I - x), 3).Value
        Max_1 = Cells(I, 3).Value
        Growth = Max_1 - Min_1
        If Min_1 = 0 Then
        Percent_Growth = 0
        Else
        Percent_Growth = ((Max_1 - Min_1) / Min_1)
        End If
        Range("J" & Summary_table_row).Value = Growth
        Range("K" & Summary_table_row).Value = Percent_Growth
        Summary_table_row = Summary_table_row + 1
        ticker_total = 0
        x = 0
    Else
    
    'This section gets executed when the ticker value is the same
        ticker_total = ticker_total + Cells(I, 7).Value
        x = x + 1
    End If

Next I


'The below section of code is to apply conditional formatting for negative nad positive change in stock prices

Set Cond1 = rg.FormatConditions.Add(xlCellValue, xlGreater, "=0.00")
Set Cond2 = rg.FormatConditions.Add(xlCellValue, xlLess, "=0.00")
With Cond1
.Interior.Color = vbGreen
.Font.Color = vbBlack
End With
With Cond2
.Interior.Color = vbRed
.Font.Color = vbBlack
End With

'This section of code is to get the maximum and minimum value from the yearly change column and to get he maximum total stock volume

Max = Application.WorksheetFunction.Max(Columns("k"))
Min = Application.WorksheetFunction.Min(Columns("k"))
Max_Total = Application.WorksheetFunction.Max(Columns("L"))


Range("Q2").Value = Max
Range("Q3").Value = Min
Range("Q4").Value = Max_Total

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"
Range("I1:Q1").Font.Bold = True

End Sub




