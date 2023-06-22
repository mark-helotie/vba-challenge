Sub StockAnalysis()

' Define Variables
Dim CurrRow, OpenPrice, SummChartRow, MaxRow, MinRow, VolRow As Integer
Dim MaxInc, MinInc As Double
Dim VolMax, CumulVolume As LongLong
    
' Loop thru all the worksheets ---- HERE WE GO!!!
For Each ws In Worksheets
        
    ' Initialize Some variables
    SummChartRow = 2                   ' Summary Chart starts on row 2
    CumulVolume = 0                    ' Start from scratch for first ticker
    OpenPrice = ws.Cells(2, 3).Value   ' Set opening price for first ticker

    ' Find out how many rows on this page
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ' Format Column Headers for Summary data
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Range("I1:Q1").Font.Bold = True
    ws.Range("O2:O4").Font.Bold = True
 
    ' Format Cells for Summary Data
    ws.Range("J:J").NumberFormat = "##0.00"
    ws.Range("K:K").Style = "Percent"
    ws.Range("K:K").NumberFormat = "##0.00%"
    ws.Range("L:L").NumberFormat = "#,##0"
    ws.Range("Q2:Q3").Style = "Percent"
    ws.Range("Q2:Q3").NumberFormat = "##0.00%"
    ws.Range("Q4").NumberFormat = "#,##0"
    ws.Columns("I:Q").AutoFit
        
    ' Start looping thru all the rows in this sheet
    For CurrRow = 2 To LastRow
        
        ' Confirm we are still within the same ticker
        If ws.Cells(CurrRow + 1, 1).Value = ws.Cells(CurrRow, 1).Value Then
            CumulVolume = CumulVolume + ws.Cells(CurrRow, 7).Value

        Else    ' We have a NEW ticker coming next!!!
                ' Put current ticker into the Summary Chart
            ws.Cells(SummChartRow, 9).Value = ws.Cells(CurrRow, 1).Value
            ' Compute yearly change and put into Summary Chart
            ws.Cells(SummChartRow, 10).Value = ws.Cells(CurrRow, 6).Value - OpenPrice
                If ws.Cells(SummChartRow, 10).Value >= 0 Then
                    ws.Cells(SummChartRow, 10).Interior.ColorIndex = 4     ' positive = GREEN
                Else
                    ws.Cells(SummChartRow, 10).Interior.ColorIndex = 3     ' negative = RED
                End If
            ' Compute percent change and put into Summary Chart
            ws.Cells(SummChartRow, 11).Value = (ws.Cells(CurrRow, 6) - OpenPrice) / OpenPrice
                If ws.Cells(SummChartRow, 11).Value >= 0 Then
                   ws.Cells(SummChartRow, 11).Interior.ColorIndex = 4     ' positive = GREEN
                Else
                   ws.Cells(SummChartRow, 11).Interior.ColorIndex = 3     ' negative = RED
                End If
            ' Compute Cumulative Volume and put into Summary Chart
            CumulVolume = CumulVolume + ws.Cells(CurrRow, 7).Value
            ws.Cells(SummChartRow, 12).Value = CumulVolume
            CumulVolume = 0      ' reset for the next ticker
            OpenPrice = ws.Cells(CurrRow + 1, 3).Value   ' set opening price for next ticker
            SummChartRow = SummChartRow + 1              ' done with this summary row
                
        End If
            
    Next CurrRow
    
    ' We have processed ALL the worksheets... now we do the Greatest chart.

    ' Find max/min values in the summary table
    MaxInc = Application.WorksheetFunction.Max(ws.Range("K:K"))
    MinInc = Application.WorksheetFunction.Min(ws.Range("K:K"))
    VolMax = Application.WorksheetFunction.Max(ws.Range("L:L"))
    
    ' Find the row number of each of the above variables
    MaxRow = Application.WorksheetFunction.Match(MaxInc, ws.Range("K:K"), 0)
    MinRow = Application.WorksheetFunction.Match(MinInc, ws.Range("K:K"), 0)
    VolRow = Application.WorksheetFunction.Match(VolMax, ws.Range("L:L"), 0)
      
    ' Populate Greatest table with values found above for each min/max
    ws.Range("Q2").Value = MaxInc
    ws.Range("Q3").Value = MinInc
    ws.Range("Q4").Value = VolMax
    
    ' Using the row numbers found above, populate the ticker symbols
    ws.Range("P2").Value = ws.Cells(MaxRow, 9).Value
    ws.Range("P3").Value = ws.Cells(MinRow, 9).Value
    ws.Range("P4").Value = ws.Cells(VolRow, 9).Value
    
    ' Autofit the columns to make it pretty :)
    ws.Columns("I:Q").AutoFit
    
Next ws
    
End Sub
