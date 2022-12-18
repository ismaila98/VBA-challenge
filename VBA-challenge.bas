''Attribute VB_Name = "Module1"

Sub stock_market()
Dim last_row As Long
Dim column As Long
Dim closing_price As Double
Dim opening_price As Double
Dim yearly_change As Double
Dim percentage_change As Double
Dim increment As Long

increment = 0
column = 1

'finds last row of the active sheet
last_row = Cells(Rows.Count, 1).End(xlUp).Row
    'looping through all the rows
    For i = 2 To last_row
        'when it finds a different value in the ticker row
        If Cells(i + 1, column).Value <> Cells(i, column).Value Then
            increment = increment + 1
            'returns ticker
            Cells(1 + increment, 9) = Cells(i, 1).Value
            'returns the difference between the end year closing price and the beginning of the year open price
            closing_price = Cells(i, 6).Value
            opening_price = Cells(i - 250, 3).Value
            yearly_change = closing_price - opening_price
            Cells(1 + increment, 10).Value = yearly_change
            'shows percentage change in price
            percentage_change = (yearly_change / opening_price) * 100
            Cells(1 + increment, 11).Value = percentage_change
            'shows total stock volume
            Cells(1 + increment, 12).Formula = "=SUM(" & Range(Cells(i - 250, 7), Cells(i, 7)).Address(False, False) & ")"
            If yearly_change > 0 Then
                'formats positive change in green
                Cells(1 + increment, 10).Interior.ColorIndex = 4
            ElseIf yearly_change < 0 Then
                'formats negative change in red
                Cells(1 + increment, 10).Interior.ColorIndex = 3
                Else
            End If
            Else
        End If
    Next i
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percentage Change"
Range("L1").Value = "Total Stock Volume"
End Sub
