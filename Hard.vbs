Sub hard()

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
                ws.Cells(TickerCounter, 12).Value = 0
                
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
    
    'autofit all columns
    ws.Columns("A:N").AutoFit
    

    'calcluates the Greatest % Increase, Greatest % Decrease, and Greatest Total Volume

    'Column and Row Titles

    ws.Cells(1, 18).Value = "Ticker"
    ws.Cells(1, 19).Value = "Value"
    ws.Cells(2, 17).Value = "Greatest % Increase"
    ws.Cells(3, 17).Value = "Greatest % Decrease"
    ws.Cells(4, 17).Value = "Greatest Total Volume"


    'calculate new last row for summary column
    Dim SummaryLastRow As Long
    
    SummaryLastRow = ws.Cells(ws.Rows.Count, 10).End(xlUp).Row
    'MsgBox ("last row in summary is " & SummaryLastRow)
    'search summary column for Greatest % Increase
    
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestTotalVolume As Double
    Dim IncreaseCounter As Integer
    Dim DecreaseCounter As Integer
    Dim VolumeCounter As Integer
        
    GreatestIncrease = 0
    GreatestDecrease = 0
    GreatestTotalVolume = 0
    IncreaseCounter = 0
    DecreaseCounter = 0
    VolumeCounter = 0
    
    For i = 2 To SummaryLastRow
        
        If ws.Cells(i, 12).Value > GreatestIncrease Then
            GreatestIncrease = ws.Cells(i, 12).Value
            IncreaseCounter = i
        End If
        
        If ws.Cells(i, 12).Value < GreatestDecrease Then
            GreatestDecrease = ws.Cells(i, 12).Value
            DecreaseCounter = i
        End If
        
        If ws.Cells(i, 13).Value > GreatestTotalVolume Then
            GreatestTotalVolume = ws.Cells(i, 13).Value
            VolumeCounter = i
        End If
        
    Next i
    
    ws.Cells(2, 18).Value = ws.Cells(IncreaseCounter, 10).Value
    ws.Cells(2, 19).Value = Format(GreatestIncrease, "Percent")
    ws.Cells(3, 18).Value = ws.Cells(DecreaseCounter, 10).Value
    ws.Cells(3, 19).Value = Format(GreatestDecrease, "Percent")
    ws.Cells(4, 18).Value = ws.Cells(VolumeCounter, 10).Value
    ws.Cells(4, 19).Value = GreatestTotalVolume
    
    ws.Columns("Q:S").AutoFit
    
Next ws


End Sub

