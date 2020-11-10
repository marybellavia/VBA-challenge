Sub StockMarket()

    ' setting up variables
    Dim Ticker As String
    Dim TotalVolume As Double
    Dim TickerRow As Double
    Dim YearlyOpen As Double
    Dim YearlyClose As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    
    ' most outer loop to loop through sheet
    For Each wsheet In Worksheets
    
    ' setting up variables need inside outer loop
    Ticker = wsheet.Cells(2, 1).Value
    TotalVolume = 0
    TickerRow = 2
    YearlyOpen = 2
    
    
    ' creating table
    wsheet.Cells(1, 9).Value = "Ticker"
    wsheet.Cells(1, 10).Value = "Yearly Change"
    wsheet.Cells(1, 11).Value = "Percent Change"
    wsheet.Cells(1, 12).Value = "Total Stock Volume"
    
    ' setting up an index loop to iterate through all the rows, i represents the row number
    For i = 2 To wsheet.Cells(Rows.Count, 1).End(xlUp).Row
    
        ' conditional testing if the ticker at the current index matches the previous ticker
        If wsheet.Cells(i, 1).Value = Ticker Then
            
            ' printing the Ticker value to the correct row
            wsheet.Cells(TickerRow, 9).Value = Ticker
            ' adding to the total volume
            TotalVolume = TotalVolume + wsheet.Cells(i, 7)
            ' printing to the Total Stock Volume cell in the correct row
            wsheet.Cells(TickerRow, 12) = TotalVolume
            
            ' conditional to check what yearly open is so that it will set the variable correctly in the iteration
            If YearlyOpen = 1 Then
                ' this is if the yearlyopen is for any group after the 1st
                YearlyOpen = wsheet.Cells((i - 1), 3).Value
            ElseIf YearlyOpen = 2 Then
                ' this is for if it is the first yearlyopen
                YearlyOpen = wsheet.Cells(i, 3).Value
            End If
            
        ' if the ticker is new
        ElseIf wsheet.Cells(i, 1).Value <> Ticker Then
            
            ' resetting the volume to 0
            TotalVolume = 0
            ' setting the next ticker symbol to the Ticker variable
            Ticker = wsheet.Cells(i, 1).Value
            ' setting the yearly close
            YearlyClose = wsheet.Cells((i - 1), 6).Value
            ' calculating and printing the yearly change based on yearlyopen and close
            YearlyChange = YearlyClose - YearlyOpen
            wsheet.Cells(TickerRow, 10).Value = YearlyChange
            
            ' conditional to set color of the yearly change cell
            If YearlyChange >= 0 Then
                wsheet.Cells(TickerRow, 10).Interior.ColorIndex = 4
            Else
                wsheet.Cells(TickerRow, 10).Interior.ColorIndex = 3
            End If
            
            'calculating percent change based on yearlychange
            If YearlyOpen = 0 Then
                PercentChange = 0
            Else
                PercentChange = (YearlyChange / YearlyOpen)
            End If
            ' printing the percent change
            wsheet.Cells(TickerRow, 11).Value = FormatPercent(PercentChange)
            ' setting the Yearly back at one for the other part of the if statement
            YearlyOpen = 1
            ' upping the TickerRow so that the data will go to the next row when printing
            TickerRow = TickerRow + 1
        
        End If
    Next i
    
    ' setting bonus activity variables
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestTotalVolume As Double
    
    GreatestIncrease = 0
    GreatestDecrease = 0
    GreatestTotalVolume = 0
    
    ' creating table
    wsheet.Cells(2, 14).Value = "Greatest % Increase"
    wsheet.Cells(3, 14).Value = "Greatest % Decrease"
    wsheet.Cells(4, 14).Value = "Greatest Total Volume"
    wsheet.Cells(1, 15).Value = "Ticker"
    wsheet.Cells(1, 16).Value = "Value"
    
    
    ' loop for going through results
    For b = 2 To wsheet.Cells(Rows.Count, 1).End(xlUp).Row
        ' conditional for increase and printing result
        If wsheet.Cells(b, 11).Value > GreatestIncrease Then
            GreatestIncrease = wsheet.Cells(b, 11).Value
            wsheet.Cells(2, 15).Value = wsheet.Cells(b, 9).Value
            wsheet.Cells(2, 16).Value = FormatPercent(GreatestIncrease)
        End If
        
        ' conditional for decrease and printing result
        If wsheet.Cells(b, 11).Value < GreatestDecrease Then
            GreatestDecrease = wsheet.Cells(b, 11).Value
            wsheet.Cells(3, 15).Value = wsheet.Cells(b, 9).Value
            wsheet.Cells(3, 16).Value = FormatPercent(GreatestDecrease)
        End If
        
        ' conditional for greatest total value and printing result
        If wsheet.Cells(b, 12).Value >= GreatestTotalVolume Then
            GreatestTotalVolume = wsheet.Cells(b, 12).Value
            wsheet.Cells(4, 15).Value = wsheet.Cells(b, 9).Value
            wsheet.Cells(4, 16).Value = GreatestTotalVolume
        End If
    
    Next b
    
    Next
    
    MsgBox ("Success!")
End Sub