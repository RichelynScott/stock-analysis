# Stock Analysis
Week 2- VBA Introduction

## 1. Overview of Project
**CLASS ASSIGNMENT**

Stock Analysis for 12 Stocks with tickers of AY, CSIG, DQ, ENPH, FSLR, HASI, JKS, RUN, SEDG, SPWR, TERP, VSLR. These are all "green stocks" that a friend's parents wanted some information on before investing. Visual Basic Application (VBA) was used within an excel file that had data for each ticker of stock prices-open and close values and total volume values that we used when running the VBA code to perform analysis. Main outcome was to output a Total Daily Value for each Ticker value, or stock, and the annual return for each of the 12 in this data set. With the analysis we can see which stocks in this set for those years performed the best to worst.

# 2. Results
Refactoring the Code

The objective of refactoring the code was to code to loop through all the data of a given year and return the total daily volume of each stock/ticker and annual return for each as well. In order to make the code more efficient, I needed to switch the nesting order of my for loops. To do this, I created a 4 different arrays; tickers (each stock organized by their 12 respective ticker values) , tickerVolumes (to calculate and add all the volumes for each ticker), tickerStartingPrices (to help calculate return), and tickerEndingPrices (to help calculate return). The tickers array was used to establish the ticker symbol of a stock. I matched the other three arrays with the tickers array by using a variable called the tickerIndex.

###### Refactored Code

    Sub VBA_Challenge()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
    
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
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    Dim tickerIndex As Single
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single

    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
    tickerVolumes(i) = 0
    Next i
    
    ''2b) Loop over all the rows in the spreadsheet.
    For j = 2 To RowCount

        '3a) Increase volume for current ticker
    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
    If Cells(j - 1, 1).Value <> tickers(tickerIndex) Then
         tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
            
            
        'End If
    End If
    
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
    If Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
            

            '3d Increase the tickerIndex.
        tickerIndex = tickerIndex + 1
            
        'End If
        End If
        
    Next j
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        tickerIndex = i
        Cells(i + 4, 1).Value = tickers(tickerIndex)
        Cells(i + 4, 2).Value = tickerVolumes(tickerIndex)
        Cells(i + 4, 3).Value = (tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex)) - 1
        
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

    End Sub

###### Original Code

    '2) Initialize array of all tickers

    Dim tickers(12) As String

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

  '3a) Initialize variables for starting price and ending price

    Dim startingPrice As Double
    Dim endingPrice As Double

  '3b) Activate data worksheet

    Worksheets(yearValue).Activate

  '3c) Get the number of rows to loop over

    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

  '4) Loop through tickers

    For i = 0 To 11
    ticker = tickers(i)
    TotalVolume = 0
    Worksheets(yearValue).Activate

  '5) loop through rows in the data
        
  For j = 2 To RowCount

    '5a) Get total volume for current ticker

    If Cells(j, 2).Value = ticker Then

        'increase totalVolume by the value in the current row
        TotalVolume = TotalVolume + Cells(j, 8).Value

  End If

        '5b) get starting price for current ticker

    If Cells(j - 1, 2).Value <> ticker And Cells(j, 2).Value = ticker Then
        'set starting price
        startingPrice = Cells(j, 6).Value

    End If

        '5c) get ending price for current ticker
        
        If Cells(j + 1, 2).Value <> ticker And Cells(j, 2).Value = ticker Then
        'set ending price
        endingPrice = Cells(j, 6).Value

    End If

    Next j
'6) Output data for current ticker

    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = TotalVolume
    Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

 Next i


This variable allowed me to assign the tickerVolumes, tickerStartingPrices, and tickerEndingPrices to each ticker symbol before interating through the data set. By doing it this way, the analysis would be completed much faster than using the nested for loop for earlier.

## Elapsed Time to Run for each year (2017 & 2018)


2017 Analysis Elpsed Time PNG- https://github.com/RichelynScott/stock-analysis/blob/main/VBA_Challenge_2017.png

2018 Analysis Elapsed Time PNG- https://github.com/RichelynScott/stock-analysis/blob/main/VBA_Challenge_2018.png

## Summary of Refactoring

The pros and cons to refactoring lean to it being more beneficial than not, but it really depends on the specific project and goals. The benefits of refactoring are that you are already working with a VBA script that works from an original code and can tweak and modify to make the code more efficient and potentially run faster. The disadvantages of Refactoring are that you can possible brink, or make your code unusable due to an error in the new code, you can also spend a lot more time fixing issues with your new code if you get the syntax wrong or miss a variable assignment, etc. As stated before, it would depend on the project and its outcome goals whether or not refactoring would be a worthy endeavor.
