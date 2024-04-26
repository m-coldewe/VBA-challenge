Sub StockMarketMan()

'set variable for worksheets
Dim WS As Worksheet


For Each WS In Worksheets
   
    'Create column headers for Summary Table
    WS.Range("I1").Value = "Ticker"
    WS.Range("J1").Value = "Yearly Change"
    WS.Range("K1").Value = "Percent Change"
    WS.Range("L1").Value = "Total Stock Volume"
   
    WS.Columns("i:i").EntireColumn.AutoFit
    WS.Columns("j:j").EntireColumn.AutoFit
    WS.Columns("k:k").EntireColumn.AutoFit
    WS.Columns("l:l").EntireColumn.AutoFit
    
    WS.Range("P1").Value = "Ticker"
    WS.Range("Q1").Value = "Value"
 

    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2


    'loop through all ticker values
    For x = 2 To WS.UsedRange.Rows.Count
        'set initial variable for holding tickerID
        Dim tickerID As String

        'set initial variable for holding tickerVolume
        Dim tickerVolume As Double

        'set tickerVolume to 0
        tickerVolume = 0

        'set initial variable for holding tickeropeningamt
        Dim tickerOpenAmt As Double

        'set initial variable for holding tickerclosingamt
        Dim tickerCloseAmt As Double

        'set initial variable for Yearly Change
        Dim yearChange As Double

        'set initial variable for holding Percent Change
        Dim percentChange As Double


        If WS.Cells(x - 1, 1).Value <> WS.Cells(x, 1).Value Then

            tickerOpenAmt = WS.Cells(x, 3).Value
    
            End If
    
    
        'if we are not in the same tickerID
        If WS.Cells(x + 1, 1).Value <> WS.Cells(x, 1).Value Then

            'set the tickerID
            tickerID = WS.Cells(x, 1).Value
    
            tickerCloseAmt = WS.Cells(x, 6).Value
    
            'subtract opening amount from closing amount
            yearChange = tickerCloseAmt - tickerOpenAmt
    
            'calculate percent change
            percentChange = yearChange / tickerOpenAmt * 100
    
            'add to the ticker volume
            tickerVolume = tickerVolume + WS.Cells(x, 7).Value
    
            'print tickerId to the Summary Table
            WS.Range("I" & Summary_Table_Row).Value = tickerID
    
            'print yearChange to Sumarry Table
            WS.Range("J" & Summary_Table_Row).Value = yearChange
    
            'print Percent Change to Summary Table
            WS.Range("K" & Summary_Table_Row).Value = percentChange
            
            'print ticker volume to Summary Table
            WS.Range("L" & Summary_Table_Row).Value = tickerVolume
    
            'add 1 to the Summary Row Table
            Summary_Table_Row = Summary_Table_Row + 1
    
            tickerVolume = 0
    
    
            'if the cell immediately following a row is the same ticker
        Else
    
            'add to ticker volume
            tickerVolume = tickerVolume + WS.Cells(x, 7).Value
            
    
        End If


    Next x

    Dim y As Long
    Last_Row = WS.Cells(WS.Rows.Count, 10).End(xlUp).Row
    For y = 2 To Last_Row
        'use conditional formating to color Yearly Change
        If (WS.Cells(y, 10).Value > 0) Then
            WS.Cells(y, 10).Interior.Color = RGB(0, 255, 0)

        ElseIf (WS.Cells(y, 10).Value < 0) Then
            WS.Cells(y, 10).Interior.Color = RGB(255, 0, 0)
        
        End If

        'use conditional formating to color Percent Change
        If (WS.Cells(y, 11).Value > 0) Then
            WS.Cells(y, 11).Interior.Color = RGB(0, 255, 0)

        ElseIf (WS.Cells(y, 11).Value < 0) Then
            WS.Cells(y, 11).Interior.Color = RGB(255, 0, 0)

        End If
        
    Next y

    'Create column headers for percentages table
    WS.Range("P1").Value = "Ticker"
    WS.Range("Q1").Value = "Value"
   
    WS.Range("O2").Value = "Greatest % Increase"
    WS.Range("O3").Value = "Greatest % Decrease"
    WS.Range("O4").Value = "Greatest Total Volume"
    
    'delcare max and min variables
    Dim maxPercent As Double
    Dim minPercent As Double
    Dim maxVolume As Double


    Dim tempIndex As Double
    Dim tempIndex2 As Double
    Dim tempIndex3 As Double

    'calculate max and min
    maxPercent = Application.WorksheetFunction.Max(WS.Range("K:K"))
    maxVolume = Application.WorksheetFunction.Max(WS.Range("L:L"))
    minPercent = Application.WorksheetFunction.Min(WS.Range("K:k"))

    tempIndex = WorksheetFunction.Match(maxPercent, WS.Range("K:K"), 0)
    WS.Range("P2").Value = WS.Cells(tempIndex, 9).Value

    tempIndex2 = WorksheetFunction.Match(minPercent, WS.Range("K:K"), 0)
    WS.Range("P3").Value = WS.Cells(tempIndex2, 9).Value

    tempIndex3 = WorksheetFunction.Match(maxVolume, WS.Range("L:L"), 0)
    WS.Range("P4").Value = WS.Cells(tempIndex3, 9).Value

    'assign values for max and min
    WS.Range("Q2").Value = maxPercent
    WS.Range("Q3").Value = minPercent
    WS.Range("Q4").Value = maxVolume

    WS.Columns("o:o").EntireColumn.AutoFit
    WS.Columns("Q:Q").EntireColumn.AutoFit
    
    
Next

End Sub