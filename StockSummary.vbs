Sub StockSummary():
    'Declare variables
    Dim tickerSym As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim tickerVol As Double
    Dim Row As Long
    Dim sumRow As Long
    
    'Decalre Bonus variables
    Dim GrtPctIncrease As Double
    Dim GrtPctIncTicker As String
    Dim GrtPctDecrease As Double
    Dim GrtPctDecTicker As String
    Dim GrtTotVolume As Double
    Dim GrtTotVolTicker As String

    'Use "ws" to make changes in all worksheets
    For Each ws In Worksheets
        'Set header columns for a summary of annual data
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        'Initialize row and annual summary row variables
        Row = 2
        sumRow = 2
        
        'Initialize bonus variables
        GrtPctIncrease = 0
        GrtPctDecrease = 0
        GrtTotVolume = 0
        
        'Use a while loop to keep moving through every row on each annual data sheet
        While Not IsEmpty(ws.Cells(Row, 2).Value)
            tickerSym = ws.Cells(Row, 1).Value  'set ticker symbol value
            openPrice = CDbl(ws.Cells(Row, 3).Value)    'set opening price
            tickerVol = 0   'Initialize ticker volume variable
            
            'While loop to add up all of the ticker volume for a single ticker symbol
            While ws.Cells(Row, 1).Value = tickerSym
                tickerVol = tickerVol + CDbl(ws.Cells(Row, 7).Value)
                Row = Row + 1
            Wend
            
            closePrice = CDbl(ws.Cells(Row - 1, 6).Value)   'set close price at the end of the year

            'Summarize the annual data for the current ticker symbol
            ws.Cells(sumRow, 9).Value = tickerSym
            ws.Cells(sumRow, 10).Value = (closePrice - openPrice)
            
            'Format the change in closing price based on gain or loss
            If ws.Cells(sumRow, 10).Value > 0 Then
                ws.Cells(sumRow, 10).Interior.ColorIndex = 4
            ElseIf ws.Cells(sumRow, 10).Value < 0 Then
                ws.Cells(sumRow, 10).Interior.ColorIndex = 3
            End If
            
            'Calculate the percent change; If statement to avoid overflow error
            If openPrice > 0 Then
                percentchange = ((closePrice - openPrice) / openPrice)
            Else: percentchange = 0
            End If
            
            'Track the greatest percent increase/decrease and greatest total volume for each year
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
            sumRow = sumRow + 1     'Move to the next summary row
            
        Wend
        
        'Print out the greatest pecent increase/decrease and greatest total volume for each year
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
        ws.Range("Q2:Q3").NumberFormat = "0.00%"  'format the percent increase/decrease
        ws.Range("Q4").NumberFormat = "##0.0000E+0" 'format the greatest total volume in scientific notation
        ws.Columns("I:Q").AutoFit   'format column widths to fit summary data and greatest increase/decreas/volume summary
    Next ws

End Sub