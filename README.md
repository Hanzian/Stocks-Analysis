# **Stocks-Analysis**

  ## **Overview of the project**
  
   ### **Background**
   
Steve would like help from his friend Data Analyst for a relevant analysis of his parents' investment in green energy stocks. For this we used VBA to automate the analysis and allow Steve to use it in the future.
We started by creating a worksheet to analyse the list of stocks, their daily volume and annual returns (the percentage difference between the price at the end of the year and that of the beginning of the year).
Also we used conditional formatting to facilitate analysis.
Finally, we inserted a code allowing us to calculate the execution time of a code.
    
   ### **Purpose of  project**
        
Steve wants to expand the dataset to include the entire stock market over the last few years. Thus, the purpose of this project is to refactor the code to allow for a large dataset to be analyzed in an efficient amount of time. The refactored code will loop through all the data only one time to output the total volume and yearly return.  

## **Results**

  ### **Refactoring**
  
To make my code more efficient, I created 3 new arrays:

- tickerVolumes(12) to hold volume
- tickerStartingPrices(12) to hold starting price
- tickerEndingPrices(12) to hold ending price

The above 3 arrays store performance data for each stock when a for loop runs analysis on them. The tickers array that I created in the original establishes a ticker symbol that can be called on for each stock.

Matching the 3 performance arrays with the ticker array is done by using a variable called the tickerIndex.

Now that I have created these arrays, I can use Nested For Loops and variables to loop through the data and complete the analysis. 

   #### **Refacored Code**
   ```
   Sub AllStocksAnalysisRefactored()
    
    
        Dim startTime As Single
        Dim endTime  As Single

        YearValue = InputBox("What year would you like to run the analysis on?")
    
        startTime = Timer
        
        'Format the output sheet on All Stocks Analysis worksheet
        Worksheets("All Stocks Analysis").Activate
        
        Range("A1").Value = "All Stocks (" + YearValue + ")"
        
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
    Worksheets(YearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
        
        Dim tickerIndex As Integer
    'Set equal to zero
    
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
        For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        
                If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
        
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        
                End If

    
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex
        
                If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
        
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        
                End If
            
        
            '3d Increase the tickerIndex.
                
                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
                tickerIndex = tickerIndex + 1
                
                End If

            
        'End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
         Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
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
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (YearValue)
   
   End Sub 
   ```
   #### **Original Code**
   ```
    Sub AllStocksAnalysis()
    
     Dim startTime As Single
    Dim endTime  As Single
    
 
        YearValue = InputBox("What year would you like to run the analysis on?")
        
       'set time start
       
        starterTime = Timer
        
        'Activate "All Stocks Analysis"
        
        Worksheets("All Stocks Analysis").Activate
                
        'Title analysis
        
        Range("A1").Value = "All Stoks (" + YearValue + ")"
        
        'Create a header row
    
    Cells(3, 1).Value = "Ticker"
    
    Cells(3, 2).Value = "Total Daily Volume"
    
    Cells(3, 3).Value = "Return"
    
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
   
   Worksheets(YearValue).Activate
   
   '3c) Find the number of rows to loop over.
        
        rowStart = 2
        'DELETE: rowEnd = 3013
        'rowCount code taken from https://stackoverflow.com/questions/18088729/row-count-where-data-exists
        RowCount = Cells(Rows.Count, "A").End(xlUp).Row
        
        
        '4) Loop through tickers
   
   For i = 0 To 11
   
       ticker = tickers(i)
       
       totalVolume = 0
       
       '5) loop through rows in the data
       
       Worksheets(YearValue).Activate
       
       For j = rowStart To RowCount
       
    '5a) Get total volume for current ticker
    
    If Cells(j, 1).Value = ticker Then
    
    
        totalVolume = totalVolume + Cells(j, 8).Value
        
        End If
        
           
    '5b) get starting price for current ticker
    
           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
           
           startingPrice = Cells(j, 6).Value
           
           End If
           
    '5c) get ending price for current ticker
    
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            
            endingPrice = Cells(j, 6).Value
            
            End If

       Next j
       
       '6) Output data for current ticker
    
    Worksheets("All Stocks Analysis").Activate
    
    Cells(4 + i, 1).Value = ticker
    
    Cells(4 + i, 2).Value = totalVolume
    
    Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
    
    
    Next i
 
  
        endTime = Timer
  
            MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (YearValue)
            Worksheets("All Stocks Analysis").Activate
            
            
     End Sub
     ```
     ### **Stocks Performances**
 


    
