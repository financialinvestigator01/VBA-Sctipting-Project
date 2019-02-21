Sub easy()

'Create a script that will loop through each year of stock data and grab the total amount of volume each stock had over the year.

    'Looping for all worksheets
    
Dim ws As Worksheet

For Each ws In Worksheets


    'Find the last row and column in each worksheet
    
    Dim LastRow As Long
    Dim LastColumn As Long
               
    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    LastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    'MsgBox ("last Row: " & LastRow & " & Last Column: " & LastColumn)
    
    'end find last row and column
    
    'label Ticker and Total Stock Volume columns
    
    ws.Cells(1, 10).Value = "Ticker"
    ws.Cells(1, 11).Value = "Total Stock Volume"
    
    'end column labels
       

'loop "ticker" column thru each new symbol and keep a running total of its volume

    Dim i As Long
    Dim TickerCounter As Integer
    Dim TickerSymbolVolume As Double
    
    'read range of cells in first colum for each different ticker symbol
    
    TickerSymbolVolume = 0
    TickerCounter = 1
    
    For i = 2 To LastRow
        
        If ((ws.Cells(i, 1).Value) = (ws.Cells(i + 1, 1).Value)) Then
            'sum the total volume of each ticker symbol
            TickerSymbolVolume = TickerSymbolVolume + ws.Cells(i, 7).Value
            
        Else
            TickerSymbolVolume = TickerSymbolVolume + ws.Cells(i, 7).Value
            TickerCounter = TickerCounter + 1
            ws.Cells(TickerCounter, 10).Value = ws.Cells(i, 1).Value
            ws.Cells(TickerCounter, 11).Value = TickerSymbolVolume
            TickerSymbolVolume = 0
        
        End If
             
    Next i

Next ws


End Sub








