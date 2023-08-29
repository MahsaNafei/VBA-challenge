Sub Multiple_year_stock_data():

For Each ws In Worksheets

lastrow = ws.Cells(Rows.count, 1).End(xlUp).row

Dim Volcount As Double
Dim ticker As String
Dim row As Double
Dim openvalue As Double
Dim closevalue As Double
Dim first As Double
Dim count As Double
Dim yearlychange As Double
Dim percentchange As Double


ws.Range("I1,P1").Value = "Ticker"
ws.Range("J1").Value = "Yeary Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("Q1").Value = "Value"

row = 2

For i = 2 To lastrow
 If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    ticker = ws.Cells(i, 1).Value
    ws.Range("I" & row).Value = ticker
    ws.Range("L" & row).Value = Volcount
    ws.Range("J" & row).Value = yearlychange
    ws.Range("K" & row).Value = percentchange
    ws.Range("J" & row).Style = "Currency"
    ws.Range("K" & row).NumberFormat = "0.00%"
   
    GreatestIncrease = Application.WorksheetFunction.Max(ws.Range("K2:K" & row))
    GreatestDecrease = Application.WorksheetFunction.Min(ws.Range("K2:K" & row))
    GreatestTotalVolume = Application.WorksheetFunction.Max(ws.Range("L2:L" & row))
    
   
    ws.Range("Q2").Value = GreatestIncrease
    ws.Range("Q3").Value = GreatestDecrease
    ws.Range("Q4").Value = GreatestTotalVolume
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3").NumberFormat = "0.00%"
    
    If ws.Range("K" & row) = GreatestIncrease Then
    ws.Range("P2").Value = ws.Range("I" & row).Value
    ElseIf ws.Range("K" & row) = GreatestDecrease Then
    ws.Range("P3").Value = ws.Range("I" & row).Value
    ElseIf ws.Range("L" & row) = GreatestTotalVolume Then
    ws.Range("P4").Value = ws.Range("I" & row).Value
    End If
    
    
    If yearlychange > 0 Then
    ws.Range("J" & row).Interior.ColorIndex = 4
    Else
    ws.Range("J" & row).Interior.ColorIndex = 3
    End If
    
    If percentchange > 0 Then
    ws.Range("K" & row).Interior.ColorIndex = 4
    Else
    ws.Range("K" & row).Interior.ColorIndex = 3
    End If
    
    row = row + 1
    Volcount = 0
    count = 0
    first = 0
    
 Else
    
    Volcount = Volcount + ws.Cells(i, 7).Value
    count = count + 1
    first = i - count + 1
    
    openvalue = ws.Cells(first, 3).Value
    closevalue = ws.Cells(i + 1, 6).Value
    yearlychange = closevalue - openvalue
    percentchange = ((closevalue - openvalue) / openvalue)
    
 End If

Next i

ws.Columns("A:Q").AutoFit

Next ws

End Sub




