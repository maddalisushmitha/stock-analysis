# stock-analysis with Excel

## Analysis and purpose:

Projecting stock which are profitable in years 2017 and 2018 so as to predict next years profitable stocks to help Steve and his parents make a right decision to invest by creating an optimized VBA code which is far better in terms of execution time, saving memory than the original code.

# Results:

## Stocks performance between 2017 and 2018:

1. In 2017 most of the stocks are profiting in which  DQ,ENPH,FSLR,SEDG make the highest profit.
2. As for ENPH its profiting is in decreasing pattern.

![ENPH](https://github.com/maddalisushmitha/stock-analysis/blob/main/images_for_readme/Stock-ENPH_Analysis.png)

3. When we look at the same in 2018, almost every stock is in loss state except for RUN which gained a hike of 
78.5%. This would be the best stock to buy right way.

![RUN](https://github.com/maddalisushmitha/stock-analysis/blob/main/images_for_readme/Stock-Run_Analysis1.png)

4.We can also note that the stock TERP recovered it loss a little that might make a good investment in the long future

![TERP](https://github.com/maddalisushmitha/stock-analysis/blob/main/images_for_readme/Stock-TERP_Analysis.png)

## Code and comparison of execution times of the original script and the refactored script:

### Code:

1.In the original code we are accessing  3000+ records 11 times by using two for loops where as in the refractor code it loops for 1 time reducing the run time.

#### Original code:

For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
        Worksheets(yearValue).Activate
        For j = rowStart To rowEnd
#### Refractor code:

For i = 0 To 11
        tickerVolumes(i) = 0
        Next i
    'Looping over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
 2.As we are initiating with arrays and tickerIndex the original lines of code was reduced
 
#### Original code:

   'To find total volume for a give ticker
	If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
            End If
               'To find Starting price of given ticker
            If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
                startingPrice = Cells(j, 6).Value
            End If
              'To find Ending price of given ticker 
            If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
                endingPrice = Cells(j, 6).Value
            End If

#### Refractor code:

   'To find total volume for a give ticker
	tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
            'To find Starting price of given ticker
           If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            End If
               'To find Ending price of given ticker 
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                tickerIndex = tickerIndex + 1
 End if
 
## Execution Times:

1.By refactoring  run time was reduced

### 2017

![2017](https://github.com/maddalisushmitha/stock-analysis/blob/main/images_for_readme/Runtime_2017.png)

### 2018

![2018](https://github.com/maddalisushmitha/stock-analysis/blob/main/images_for_readme/Runtime_2018.png)

## Summary:

### Advantages of refactored:

1.Saves time and memory

## Pros and Cons that are applied to refactoring the original VBA script:

1.By using arrays tickerVolumes, tickerStartingPrices and tickerEndingPrices and initializing a variable tickerIndex that acts as index to the arrays
