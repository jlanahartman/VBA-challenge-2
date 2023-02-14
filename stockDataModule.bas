Attribute VB_Name = "Module2"
Sub multiYearStockData()

    ' Declare and set worksheet
     Dim ws As Worksheet
    
    ' Loop through all stocks for one year
     For Each ws In Worksheets
    
    ' Create the headers
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"
    
    ws.Range("O2") = "Greatest % Increase"
    ws.Range("O3") = "Greatest % Decrease"
    ws.Range("O4") = "Greatest Total Volume"
    ws.Range("P1") = "Ticker"
    ws.Range("Q1") = "Value"
    
    ' Autofit the columns
    ws.Range("A1:Q4").Columns.AutoFit
    
    
    ' Declare the variables we'll need
    Dim total As LongLong   ' total stock volume

    Dim row As Long ' loop control variable that will go through all the rows in the sheet
    
    Dim yearlyChange As Double ' variable that holds yearly change for each stock in a sheet
    
    Dim percentChange As Double    ' variable that holds the percent change for each stock in a sheet
    
    Dim rowCount As Long ' variable that holds the number of rows in a sheet
    
    Dim summaryTableRow As Long ' variable that holds that holds the rows of the summary row table
    
    Dim stockStartRow As Long ' variable that holds the start of a stock's row in the sheet
    
     
    
    ' Initalize the values
    summaryTableRow = 0 ' summary table row starts at 0
    total = 0 ' total stock volume for a stock starts at 0
    yearlyChange = 0 ' yearly change starts at 0
    stockStartRow = 2 ' first stock in the sheet will populate in row 2
    
    ' get the value of last row in the sheet
    lastRow = ws.Cells(Rows.Count, "A").End(xlUp).row
    
    ' loop until we get to the end of the sheet
    For row = 2 To lastRow
    
        ' check to see if there are changes in column A (column 1)
        If ws.Cells(row + 1, "A").Value <> ws.Cells(row, "A").Value Then
        
         ' calculate the total one last time for the ticker
          total = total + ws.Cells(row, 7).Value  ' grabs the total from column G: 'volume'
        
        ' check to see if the value of the total volume is 0
        If total = 0 Then
                    ' print the results in columns I, J, K, L
                ws.Range("I" & 2 + summaryTableRow).Value = ws.Cells(row, "A").Value ' prints the stock name in column I: 'ticker'
                ws.Range("J" & 2 + summaryTableRow).Value = 0 '  prints 0 in Yearly Change column
                ws.Range("K" & 2 + summaryTableRow).Value = 0 & "%"  ' prints 0 in Percent Change column
                ws.Range("L" & 2 + summaryTableRow).Value = 0 ' prints 0 in Total stock Volume column
                
        Else
                ' find the first non zero start value
                If ws.Cells(stockStartRow, 3).Value = 0 Then
                        For findValue = stockStartRow To row
                        
                            ' check to see if the next value (or next) value does not equal 0
                            If ws.Cells(findValue, 3).Value <> 0 Then
                            stockStartRow = findValue
                            
                            ' once we have a non-zero value, break out of the loop
                            
                          Exit For
                        End If
                    
                    Next findValue
                
        End If
        
            ' Calculate the yearly change (diff in last close - first open)
        yearlyChange = ws.Cells(row, 6).Value - ws.Cells(stockStartRow, 3).Value
            ' Calculate the percent change (yearly change/ first open)
        percentChange = yearlyChange / ws.Cells(stockStartRow, 3).Value
        
         ' print the results in columns I, J, K, L
                ws.Range("I" & 2 + summaryTableRow).Value = ws.Cells(row, "A").Value ' prints the stock name in column I: 'ticker'
                ws.Range("J" & 2 + summaryTableRow).Value = yearlyChange     'prints in Yearly Change column
                ws.Range("J" & 2 + summaryTableRow).NumberFormat = "0.00"   'formats Yearly Change column
                ws.Range("K" & 2 + summaryTableRow).Value = percentChange    'prints in Percent Change column
                ws.Range("K" & 2 + summaryTableRow).NumberFormat = "0.00%"  'formats the Percent Change column
                 ws.Range("L" & 2 + summaryTableRow).Value = total    'prints in the total stock volume column
                 ws.Range("L" & 2 + summaryTableRow).NumberFormat = "#.###"   'formats the total stock volume column
           
           ' formatting for the yearlyChange column
           If yearlyChange > 0 Then
                ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 4 ' green
            ElseIf yearlyChange < 0 Then
                ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 3 ' red
            Else
                ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 0 ' white/no change
           End If
        
            End If
            
        ' reset the values of the total
        total = 0
      yearlyChange = 0
      ' move to the next row in the summary table
      summaryTableRow = summaryTableRow + 1
        
        ' if the ticker is the same
        Else
            total = total + ws.Cells(row, 7).Value  ' grabs the total from column G: 'volume'
    
        End If
        
        Next row
        
        ' after looping through the rows, find the max and min and place them in respective cells
        ws.Range("Q2").Value = "%" & WorksheetFunction.Max(ws.Range("K2:K" & lastRow)) * 100
        ws.Range("Q3").Value = "%" & WorksheetFunction.Min(ws.Range("K2:K" & lastRow)) * 100
        ws.Range("Q4").Value = "%" & WorksheetFunction.Min(ws.Range("L2:L" & lastRow)) * 100
        ws.Range("Q4").NumberFormat = "#,###"
        
        increaseNumber = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & lastRow)), ws.Range("K2:K" & lastRow), 0)
        ws.Range("P2").Value = ws.Cells(increaseNumber + 1, 9)
        
        decreaseNumber = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & lastRow)), ws.Range("K2:K" & lastRow), 0)
        ws.Range("P3").Value = ws.Cells(decreaseNumber + 1, 9)
        
        volumeNumber = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("L2:L" & lastRow)), ws.Range("L2:L" & lastRow), 0)
        ws.Range("P4").Value = ws.Cells(volumeNumber + 1, 9)
        
        ws.Columns("A:Q").AutoFit
        
    
           
        Next ws


End Sub
