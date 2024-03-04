Attribute VB_Name = "StockAnalysis"
' Start here
' Cycle through worksheets
Sub forEachWs()
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        Call StockAnalysis(ws)
        Call Bonus(ws)
    Next
End Sub


' Analyse worksheet
Sub StockAnalysis(ws As Worksheet)

    ' Workhorse variables
    Dim tickerSymbol As String
    Dim stockOpen As Double, stockClose As Double, stockVolume As Variant
    Dim stockPercentage As Double, stockPercentChange As Double
    Dim tickerCounter As Long

    ' Get last row for looping
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ' Quick display to assure user that sheets are accessing properly
    MsgBox ("Worksheet Title " & ws.Name)
    MsgBox ("Rows of data " & LastRow)
    
    ' Create new column titles for results
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ' Get first open because loop is set to "not equals"
    stockOpen = Range("C2").Value
    
    ' Get first ticker for the same reason
    tickerSymbol = ws.Range("A2").Value
    tickerCounter = 2
    
    ' Add first ticker to results column for same reason
    ws.Range("I2").Value = tickerSymbol

    ' Loop to get open and close and volume values
    ' then put the data into the new results columns
    For I = 2 To LastRow
    
        ' If next line doesn't equal last ticker
        If ws.Cells(I, 1).Value <> ws.Cells(I + 1, 1).Value Then

            ' Get the final value for the time period, ie final close value
            stockClose = ws.Cells(I, 6)
                        
            ' Add result data to new columns according to tickerCounter
            ' This is the total change result
            ws.Cells(tickerCounter, 10).Value = stockClose - stockOpen
            
            ' Change cell colour based on performance
            If ws.Cells(tickerCounter, 10).Value > 0 Then
                ws.Cells(tickerCounter, 10).Interior.ColorIndex = 4
            ElseIf ws.Cells(tickerCounter, 10).Value <= 0 Then
                ws.Cells(tickerCounter, 10).Interior.ColorIndex = 3
            End If
            
            ' Percent change result calculated and cell formatted
            ws.Cells(tickerCounter, 11).NumberFormat = "0.00%"
            ws.Cells(tickerCounter, 11).Value = (stockClose / stockOpen) - 1
            
            ' Volume calculated, displayed and variable reset
            stockVolume = stockVolume + ws.Cells(I, 7).Value
            ws.Cells(tickerCounter, 12).Value = stockVolume
            stockVolume = 0

            ' Get new stockOpen and ticker, add ticker to new results column
            stockOpen = ws.Cells(I + 1, 3).Value
            tickerSymbol = ws.Cells(I + 1, 1).Value
            tickerCounter = tickerCounter + 1
            ws.Cells(tickerCounter, 9).Value = tickerSymbol

        ' When the next line is the same ticker, add up stock volume
        Else
            stockVolume = stockVolume + ws.Cells(I, 7).Value
            
        End If
    Next I
End Sub

' Bonus round
Sub Bonus(ws As Worksheet)

    MsgBox ("Calculating bonus section")
    
    ' Make new result column titles yet again
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    ' Make new result row titles
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    ' Create variables for storage
    Dim greatestIncrease As Double, greatestDecrease As Double, greatestTotal As Variant
    Dim increaseTicker As String, decreaseTicker As String, totalTicker As String
    
    ' Last row of new columns ready for looping
    LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    ' Set initial values
    greatestIncrease = ws.Range("K2").Value
    greatestDecrease = ws.Range("K2").Value
    greatestTotal = ws.Range("L2").Value
    increaseTicker = ws.Range("I2").Value
    decreaseTicker = ws.Range("I2").Value
    totalTicker = ws.Range("I2").Value
    
    For I = 3 To LastRow
        
        ' Increase
        If ws.Cells(I, 11).Value > greatestIncrease Then
            greatestIncrease = ws.Cells(I, 11).Value
            increaseTicker = ws.Cells(I, 9).Value
        End If

        ' Decrease
        If ws.Cells(I, 11).Value < greatestDecrease Then
            greatestDecrease = ws.Cells(I, 11).Value
            decreaseTicker = ws.Cells(I, 9).Value
        End If
        
        ' Total
        If ws.Cells(I, 12).Value > greatestTotal Then
            greatestTotal = ws.Cells(I, 12).Value
            totalTicker = ws.Cells(I, 9).Value
        End If
        
    Next I
    
    ' Add values to sheet
    ws.Range("Q2").Value = greatestIncrease
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3").Value = greatestDecrease
    ws.Range("Q3").NumberFormat = "0.00%"
    ws.Range("Q4").Value = greatestTotal
    ws.Range("P2").Value = increaseTicker
    ws.Range("P3").Value = decreaseTicker
    ws.Range("P4").Value = totalTicker
    
End Sub



