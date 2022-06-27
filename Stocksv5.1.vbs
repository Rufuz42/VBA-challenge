Attribute VB_Name = "Module31"
Sub VBAHW()

Dim sheet As Worksheet ' Sets a variable to loop through multiple worksheets

For Each sheet In ThisWorkbook.Worksheets
    

    SheetName = sheet.Name 'Current work sheet's name

    StockTickerLetters = "" ' variable that holds the ticker
    Dim SumTblRow, TickerNumberCounter As Integer
    SumTblRow = 2 ' Counter for the Summary Table Rows to export data to
    TickerNumberCounter = 0 ' Counts the number of times the ticker repeats to go back and get the initial value
    Dim VolAccum As Double
    VolAccum = 0 ' Variable that accumulates stock volume
    LastRowA = sheet.Cells(Rows.Count, 1).End(xlUp).Row ' Counts the number of rows in column A
    
    ' Next 4 rows assign headers to the columns that will house the data
    sheet.Cells(1, 10).Value = "Ticker"
    sheet.Cells(1, 11).Value = "Yearly Price Change"
    sheet.Cells(1, 12).Value = "Yearly Percent Price Change"
    sheet.Cells(1, 13).Value = "Total Volume Traded"
    
    For RowA = 2 To LastRowA ' Loop through the tickers to extract summary table data
    
        If sheet.Cells(RowA + 1, 1).Value <> sheet.Cells(RowA, 1).Value Then ' Checks to see if the stock ticker changes
            StockTickerLetters = sheet.Cells(RowA, 1).Value ' Grabs the ticker letters for the current row
            VolAccum = VolAccum + sheet.Cells(RowA, 7).Value ' Adds current row's volume to accumulated volume from else loop below.
            sheet.Cells(SumTblRow, 10).Value = StockTickerLetters ' Sets the current ticker to the summary table in the right row
            sheet.Cells(SumTblRow, 13).Value = VolAccum ' Sets the current Stock Ticker Volume accumulated to column 13
            VolAccum = 0 ' Resets stock ticker accumulated volume before it starts to count again for the next ticker
            InitPrice = sheet.Cells(RowA - TickerNumberCounter, 3).Value ' Goes back up the number of rows that equals the ticker count to get year start price
            FinalPrice = sheet.Cells(RowA, 6).Value ' Since this is the last row of this ticker, column 6 has year end price.
            DeltaPrice = FinalPrice - InitPrice ' Calculates diff between final and initial price
            If DeltaPrice > 0 Then
                sheet.Cells(SumTblRow, 11).Interior.ColorIndex = 4 ' Sets cell background to green when delta price is positive
            ElseIf DeltaPrice < 0 Then
                sheet.Cells(SumTblRow, 11).Interior.ColorIndex = 3 ' Sets cell background to red when delta price is negative
            End If
            PercentPriceChange = (FinalPrice - InitPrice) / InitPrice ' calculates % change bt final and initial price
            sheet.Cells(SumTblRow, 11) = DeltaPrice  ' Sets price difference to summary table
            sheet.Cells(SumTblRow, 12) = PercentPriceChange ' Sets percent price change to summary table
            SumTblRow = SumTblRow + 1 ' Now that we have recorded to the summary table, next to go to the next row for next ticker
            TickerNumberCounter = 0 ' Resets stock ticker counter now that the ticker has changed
        Else ' Add on to total charges
            VolAccum = VolAccum + sheet.Cells(RowA, 7).Value ' Keeps adding to stock ticker volume when the ticker in the next row is the same
            TickerNumberCounter = TickerNumberCounter + 1
        End If
    
    Next RowA
    
    ' Next 3 rows count the number of rows in the fields where I spit out information
    LastRowK = sheet.Cells(Rows.Count, 11).End(xlUp).Row
    LastRowL = sheet.Cells(Rows.Count, 12).End(xlUp).Row
    LastRowM = sheet.Cells(Rows.Count, 13).End(xlUp).Row
    
    ' Next 3 rows format the numbers so they don't look goofy
    sheet.Range("K2:K" & LastRowK).NumberFormat = "$#,##0.00"
    sheet.Range("L2:L" & LastRowL).NumberFormat = "0.00%"
    sheet.Range("M2:M" & LastRowM).NumberFormat = "0,000"
    
    'Bonus Calculation
    
    'Define variables I am going to calculate, stock vol as double due to potential in size
    Dim MaxPercent, MinPercent, MaxVol As Double
    MaxPercent = 0
    MinPercent = 0
    MaxVol = 0
    
    ' Assign text to the cells that describe the values calculated
    sheet.Range("P2").Value = "Greatest % Increase"
    sheet.Range("P3").Value = "Greatest % Decrease"
    sheet.Range("P4").Value = "Greatest Total Vol"
    sheet.Range("Q1").Value = "Ticker"
    sheet.Range("R1").Value = "Value"
    
    For Row2 = 2 To LastRowL
    
        ' Define variables for max percentage and its respective ticker
        MaxPercent = WorksheetFunction.Max(sheet.Range("L2:L" & LastRowL))
        MaxPercentTicker = WorksheetFunction.Match(MaxPercent, sheet.Range("L2:L" & LastRowL), 0)
        ' Set the values determined above to specifics cells in the sheet
        sheet.Range("Q2").Value = sheet.Range("J" & MaxPercentTicker + 1).Value
        sheet.Range("R2").Value = MaxPercent
        ' Format the percentage cells as percentages
        sheet.Cells(2, 18).NumberFormat = "0.00%"
      
    Next Row2
    
    For Row3 = 2 To LastRowL
    
        ' Define variables for min percentage and its respective ticker
        MinPercent = WorksheetFunction.Min(sheet.Range("L2:L" & LastRowL))
        MinPercentTicker = WorksheetFunction.Match(MinPercent, sheet.Range("L2:L" & LastRowL), 0)
        ' Set the values determined above to specifics cells in the sheet
        sheet.Range("Q3").Value = sheet.Range("J" & MinPercentTicker + 1).Value
        sheet.Range("R3").Value = MinPercent
        ' Format the percentage cells as percentages
        sheet.Cells(3, 18).NumberFormat = "0.00%"
    
    
    Next Row3
    
    For Row4 = 2 To LastRowM
    
        ' Define variables for max volume and its respective ticker
        MaxVol = WorksheetFunction.Max(sheet.Range("M2:M" & LastRowM))
        MaxVolTicker = WorksheetFunction.Match(MaxVol, sheet.Range("M2:M" & LastRowM), 0)
        ' Set the values determined above to specifics cells in the sheet
        sheet.Range("Q4").Value = sheet.Range("J" & MaxVolTicker + 1).Value
        sheet.Range("R4").Value = MaxVol
        ' Format the stock volume as comma separated
        sheet.Cells(4, 18).NumberFormat = "0,000"
    
    Next Row4
    
Worksheets(SheetName).Range("J1:M1").Columns.AutoFit
Worksheets(SheetName).Range("P3").Columns.AutoFit
Worksheets(SheetName).Range("R4").Columns.AutoFit
Next sheet
    
    ''For Row2 = 2 To LastRowL
    ''
    ''    For Row3 = 2 To LastRowL
    ''
    ''        For Row4 = 2 To LastRowM
    ''
    ''        MaxPercent = WorksheetFunction.Max(Range("L2:L" & LastRowL))
    ''        MaxPercentTicker = WorksheetFunction.Match(MaxPercent, Range("L2:L" & LastRowL), 0)
    ''        MinPercent = WorksheetFunction.Min(Range("L2:L" & LastRowL))
    ''        MinPercentTicker = WorksheetFunction.Match(MinPercent, Range("L2:L" & LastRowL), 0)
    ''        MaxVol = WorksheetFunction.Max(Range("M2:M" & LastRowM))
    ''        MaxVolTicker = WorksheetFunction.Match(MaxVol, Range("M2:M" & LastRowM), 0)
    ''        Range("Q2").Value = Range("J" & MaxPercentTicker + 1).Value
    ''        Range("R2").Value = MaxPercent
    ''        Range("Q3").Value = Range("J" & MinPercentTicker + 1).Value
    ''        Range("R3").Value = MinPercent
    ''        Range("Q4").Value = Range("J" & MaxVolTicker + 1).Value
    ''        Range("R4").Value = MaxVol
    ''        Cells(2, 18).NumberFormat = "0.00%"
    ''        Cells(3, 18).NumberFormat = "0.00%"
    ''        Cells(4, 18).NumberFormat = "0,000"
    ''
    ''        Next Row4
    ''
    ''    Next Row3
    ''
    ''Next Row2



End Sub
