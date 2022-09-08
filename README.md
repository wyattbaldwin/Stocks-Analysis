# stock-analysis

## Overview of Project

The client, Steve, prepared a workbook with a macro enabling him to quickly analyze stock data that he has collected. Specifically, it calculates the total daily volumes for different tickers he is following and their corresponding returns. This allows him to make informed decisions when investing. He would like to continue adding on to this data, potentially even adding new tickers in the future. However, he is concerned that in the long run, the analysis will take too long. He would like us to enhance the code he has already written so that it runs faster, smoother, and for a larger dataset.

---
## Purpose

To refactor the client's macro enabling it to run faster and more efficiently for a larger dataset.

---
## Analysis and Challenges

The spreadsheet provided: [Stock Analysis](VBA_Challenge.xlsm)

### Results

The original code ran in 0.8046875 seconds and 0.78125 seconds for 2017 and 2018, respectively. Looking through the code, it was problematic in a few ways:

1. The client's code had no clear organization, which made it difficult to read. It began with declaring some variables, then moved on to formatting the output sheet, before starting the calculations, finally returning once again to formatting of the output sheet. 
```
    Dim startTime As Single
    Dim endTime As Single
    yearValue = InputBox("What year would you like to run the analysis on?")   
    startTime = Timer
  
'1) Format the output sheet on the "All Stocks Analysis" worksheet
    Worksheets("AllStocksAnalysis").Activate
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

'2) Initialize an array of all tickers

    Dim tickers(11) As String
        tickers(0) = "AY"
        tickers(1) = "CSIQ"
        tickers(2) = "DQ"
        tickers(3) = "ENPH"
        tickers(4) = "FSLR"
        tickers(5) = "HASI"
        tickers(6) = "JKS"
        tickers(7) = "RUN"
        tickers(8) = "SEDG"
        tickers(9) = "SPWR"
        tickers(10) = "TERP"
        tickers(11) = "VSLR"
```

   * The code was cleaned up to declare all values and variables upfront, followed by the necessary calculations, and ended with formatting of the output sheet. This simplifies the code, making it easier to read and follow.

```
Option Explicit
Sub AllStocksAnalysisRefactored()
    Const NUM_TICKERS As Integer = 11
    Const TICKER_COL As Integer = 1
    Const VOL_COL As Integer = 8
    Const CLOSE_COL As Integer = 6

    Dim startTime As Single
    Dim endTime  As Single
    Dim yearValue As String
    Dim tickerIndex As Integer
    Dim RowCount As Long
    Dim currentRow As Long
    Dim currentTicker As String
    
    yearValue = InputBox("What year would you like to run the analysis on?")
    startTime = Timer

    'Initialize array of all tickers
    Dim tickers(NUM_TICKERS) As String
        tickers(0) = "AY"
        tickers(1) = "CSIQ"
        tickers(2) = "DQ"
        tickers(3) = "ENPH"
        tickers(4) = "FSLR"
        tickers(5) = "HASI"
        tickers(6) = "JKS"
        tickers(7) = "RUN"
        tickers(8) = "SEDG"
        tickers(9) = "SPWR"
        tickers(10) = "TERP"
        tickers(11) = "VSLR"
```

2. The original code was littered with magic numbers, i.e. hard coded values. This creates a problem if the client wants to add tickers in the future. 

```
'4)  Loop through the tickers
     For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
        
    '5) Loop through rows in the data
        Sheets(yearValue).Activate
        For j = 2 To RowCount
        
        '5a) Find total volume for the current ticker
            If Cells(j, 1).Value = ticker Then
            
            totalVolume = totalVolume + Cells(j, 8).Value
            
            End If
        
        '5b) Find starting price for the current ticker
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            
            startingPrice = Cells(j, 6).Value
            
            End If
        
        '5c) Find the end price for the current ticker
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            
            endingPrice = Cells(j, 6).Value
            
            End If
        
        Next j
```

   * To avoid this, variables were created to define the numbers and continue to ease readability. By adding variables instead of hard coded numbers and declaring them upfront as shown in the first point, the macro can continue to run for even if the dataset is modified or increased. 

3. When the code was rearranged to a more logical flow, many redundancies were found in the original code. For example, when we moved all code relating to formatting of the output sheet to the end of the code, it was found that we only needed to activate the output worksheet once rather than twice. 

```
'1) Format the output sheet on the "All Stocks Analysis" worksheet

    Worksheets("AllStocksAnalysis").Activate
        Range("A1").Value = "All Stocks (" + yearValue + ")"
        
        'Create a header row
        Cells(3, 1).Value = "Ticker"
        Cells(3, 2).Value = "Total Daily Volume"
        Cells(3, 3).Value = "Return"
```

   * Then, at the end of the code:

```
    'Formatting
    Worksheets("AllStocksAnalysis").Activate
    Range("A3:C3").Font.Bold = True
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("A3:C3").Font.Color = vbBlue
    Range("A3:C3").Font.Italic = True
    Range("B4:B15").NumberFormat = "$#,##0.00"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit
```

   * After reorganization, we were able to remove this redundancy. We now only activate the output sheet once at the end of the code:

```
    'Formatting
    Worksheets("AllStocksAnalysis").Activate
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    'Format the table header, number formats, and column B
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "$#,##0.00"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit
```

   * Another example of redundancy from the original code is shown in the second and third IF statements (below). In sections 5b) and 5c), the original code adds an AND statement. However, this AND scenario was already addressed in section 5a). In other words, section 5a) already addresses what happens when the ticker is the same, we only need to address what happens when the ticker changes.

```
    '5a) Find total volume for the current ticker
        If Cells(j, 1).Value = ticker Then
        
        totalVolume = totalVolume + Cells(j, 8).Value
        
        End If
    
    '5b) Find starting price for the current ticker
        If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
        
        startingPrice = Cells(j, 6).Value
        
        End If
    
    '5c) Find the end price for the current ticker
        If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
        
        endingPrice = Cells(j, 6).Value
        
        End If
```

   * In the refactored code, this is removed to give the following end result:

```
    '3a) Increase volume for current ticker
    If Cells(currentRow,TICKER_COL).Value = currentTicker Then
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(currentRow, VOL_COL).Value
    Else 
        MsgBox "Error: ticker mismatch"+ CStr(currentRow) + currentTicker
        Exit Sub
    End If
    
    '3b) Check if the current row is the first row with the selected tickerIndex.
    If Cells(currentRow-1,TICKER_COL).Value <> Cells(currentRow,TICKER_COL).Value Then
        tickerstartingPrices(tickerIndex) = Cells(currentRow,CLOSE_COL).Value
    End If
    
    '3c) check if the current row is the last row with the selected ticker    
    If Cells(currentRow+1,TICKER_COL).Value <> Cells(currentRow,TICKER_COL).Value Then
        tickerendingPrices(tickerIndex) = Cells(currentRow,CLOSE_COL).Value
        tickerIndex = tickerIndex + 1
    End If
```

4. With this quick reorganization and removal of redundancies, we are able to run the code much faster than the original one. 2017 data now runs at 0.1640625 seconds and 2018 data, at 0.132815 seconds (pictured below). In addition, the refactored code is suitable to use with a larger dataset as it does not contain hard coded numbers. It relies on variables that can be modified upfront should the source data change. 

![Refactored 2017 Analysis](/Resources/VBA_Challenge_2017.png)

![Refactored 2018 Analysis](/Resources/VBA_Challenge_2018.png)

---
## Summary

Refactoring code helps make the code run faster by simplifying it and making it easier to read/understand for future use. It also helps with the macro's continuity and longevity as it allows it to become more flexible. It can however be very time-consuming to complete, especially when the code is long and complicated. It can also be risky, especially when unfamiliar with the data and the purpose of the code. By modifying the code, we may break it beyond repair. It is always best to have clear lines of communication with the client to understand the data, the code, and their intended purpose when refactoring.

In this project, we were able to refactor the client's code mainly by reorganizing and removing redundancies. By doing so, we successfully decreased the run time of the client's macro. 2017 data now runs at 0.1640625 instead of the initial 0.8046875 seconds; 2018 data now runs at 0.132815 seconds instead of 0.78125 seconds. Finally, by removing the hard coded values, the refactored macro is more flexible, requiring only simple modifications should the source dataset change significantly.
