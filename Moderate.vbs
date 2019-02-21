Sub moderate()

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
    ws.Cells(1, 11).Value = "Yearly Change"
    ws.Cells(1, 12).Value = "Percent Change"
    ws.Cells(1, 13).Value = "Total Stock Volume"
    
    'end column labels
       

'loop "ticker" column thru each new symbol and keep a running total of its volume

    Dim i As Long
    Dim TickerCounter As Integer
    Dim TickerSymbolVolume As Double
    Dim YearOpenPrice As Double
    Dim YearClosePrice As Double
        
    'read range of cells in first colum for each different ticker symbol
    
    TickerSymbolVolume = 0
    TickerCounter = 1
    
    'record stock opening price
    YearOpenPrice = ws.Cells(2, 3).Value
    
    For i = 2 To LastRow
        
        If ((ws.Cells(i, 1).Value) = (ws.Cells(i + 1, 1).Value)) Then
            'sum the total volume of each ticker symbol
            TickerSymbolVolume = TickerSymbolVolume + ws.Cells(i, 7).Value
                       
        Else
            TickerSymbolVolume = TickerSymbolVolume + ws.Cells(i, 7).Value
            TickerCounter = TickerCounter + 1
            ws.Cells(TickerCounter, 10).Value = ws.Cells(i, 1).Value
            ws.Cells(TickerCounter, 13).Value = TickerSymbolVolume
            TickerSymbolVolume = 0
            YearClosePrice = ws.Cells(i, 6).Value
            'MsgBox ("Year Open Price is: " & YearOpenPrice & " Year Close Price is: " & YearClosePrice)
            
            'calculate the diffence between the stock year open price and year closing price
            ws.Cells(TickerCounter, 11).Value = YearClosePrice - YearOpenPrice
                       
            'check for no change
            If ((YearClosePrice - YearOpenPrice) = 0) Then
                ws.Cells(TickerCounter, 12).Value = "No Change"
                
                Else
                'calculate percent change and format for percentage
                ws.Cells(TickerCounter, 12).Value = Format(((YearClosePrice - YearOpenPrice) / YearOpenPrice), "Percent")
                         
                    'format percentage cell green for positive, and red for negative
                     If (ws.Cells(TickerCounter, 12).Value >= 0) Then
                        ws.Cells(TickerCounter, 12).Interior.ColorIndex = 4
                     Else
                        ws.Cells(TickerCounter, 12).Interior.ColorIndex = 3
                     End If
                
            End If
                                                 
            'reset year open price for new ticker symbol
            YearOpenPrice = ws.Cells(1 + i, 3).Value
    

         End If
        
             
    Next i

Next ws


End Sub
