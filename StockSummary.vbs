Sub StockSummary():
    Dim tickerSym As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim tickerVol As Double
    Dim Row As Long
    Dim sumRow As Long
    
    'Bonus variables
    Dim GrtPctIncrease As Double
    Dim GrtPctIncTicker As String
    Dim GrtPctDecrease As Double
    Dim GrtPctDecTicker As String
    Dim GrtTotVolume As Double
    Dim GrtTotVolTicker As String

    For Each ws In Worksheets
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        Row = 2
        sumRow = 2
        
        'bonus variables settings
        GrtPctIncrease = 0
        GrtPctDecrease = 0
        GrtTotVolume = 0
        
        While Not IsEmpty(ws.Cells(Row, 2).Value)
            tickerSym = ws.Cells(Row, 1).Value
            openPrice = CDbl(ws.Cells(Row, 3).Value)
            tickerVol = 0
            
            While ws.Cells(Row, 1).Value = tickerSym
                tickerVol = tickerVol + CDbl(ws.Cells(Row, 7).Value)
                Row = Row + 1
            Wend
            
            closePrice = CDbl(ws.Cells(Row - 1, 6).Value)

            ws.Cells(sumRow, 9).Value = tickerSym
            ws.Cells(sumRow, 10).Value = (closePrice - openPrice)
            
            If ws.Cells(sumRow, 10).Value > 0 Then
                ws.Cells(sumRow, 10).Interior.ColorIndex = 4
            ElseIf ws.Cells(sumRow, 10).Value < 0 Then
                ws.Cells(sumRow, 10).Interior.ColorIndex = 3
            End If
            
            
            If openPrice > 0 Then
                percentchange = ((closePrice - openPrice) / openPrice)
            Else: percentchange = 0
            End If
            
            If percentchange > GrtPctIncrease Then
                GrtPctIncrease = percentchange
                GrtPctIncTicker = tickerSym
            ElseIf percentchange < GrtPctDecrease Then
                GrtPctDecrease = percentchange
                GrtPctDecTicker = tickerSym
            End If
            
            ws.Cells(sumRow, 11).Value = percentchange
            ws.Cells(sumRow, 11).Style = "Percent"
            ws.Cells(sumRow, 12).Value = tickerVol
            
            If tickerVol > GrtTotVolume Then
                GrtTotVolume = tickerVol
                GrtTotVolTicker = tickerSym
            End If
            sumRow = sumRow + 1
            
        Wend
        
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(2, 16).Value = GrtPctIncTicker
        ws.Cells(2, 17).Value = GrtPctIncrease
        ws.Cells(3, 16).Value = GrtPctDecTicker
        ws.Cells(3, 17).Value = GrtPctDecrease
        ws.Cells(4, 16).Value = GrtTotVolTicker
        ws.Cells(4, 17).Value = GrtTotVolume
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        ws.Range("Q4").NumberFormat = "##0.0000E+0"
        ws.Columns("I:Q").AutoFit
    Next ws

End Sub