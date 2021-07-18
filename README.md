# Stock Analysis Challenge 2021
## Analyzing Stock information

### Overview of Project
##### The Client for this project was interested in Stock Analysis from a group of a dozen stocks and their performance during the years of 2017 and 2018. The task at hand was to create code in VBA that would instruct Excel to loop through, analyze, format, and output stock performance in regards to volume and performance over the selected years (2017 and 2018). Our output created a basic table that included the Stocks being analyzed, their overall volume(s), and their subsequent performance. Performance results that were negative were highlighted in red and results that were positive were highlighted in green. Code was refactored from previous analysis work done for the client in order to process the results faster; The speed at which the code was processed was also recorded.
---
### Results
##### The results garnered from the Stock Analysis of years 2017 and 2018 were as follows:
The year 2017 was great for the stocks analyzed. Out of the 12 stocks analyzed for 2017, only one stock underperformed (The Stock "TERP" at -7.2%). Inverseley, the year 2018 was not great for the stocks analyzed. Out of the 12 stocks analyzed for 2018, only two stocks performed positiively (The Stocks "ENPH" and "RUN" at 81.9% and 84.0% respectively). 


<img width="445" alt="2017_Stocks_Output" src="https://user-images.githubusercontent.com/86274124/126075260-e9af7156-5ba7-49af-a9b2-71bc3cdc91c6.png">
<img width="445" alt="2018_Stocks_Output" src="https://user-images.githubusercontent.com/86274124/126075264-dc497d69-8535-4852-a4d8-24ff5584a10a.png">


For the code being run, first a basic table was formatted ni VBA followed by a delcaration of the stocks being analyzed in the following format: 

    Sub AllStocksAnalysisRefactored()
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
  
To compile the Stock Analysis from years 2017 and 2018, an index was made followed by volumes, starting prices and ending prices. A loop was than created to initialize the volumes, starting prices and ending prices for the stocks being analyzed. All information was looped through and analyzed in the stocks for 2017 and 2018 from the beginning rows to the end rows, from the starting prices to the ending prices. Volumes were compiiled and outputted to a separate worksheet along with stock performance.

    '1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
    tickerVolumes(i) = 0
    tickerStartingPrices(i) = 0
    tickerEndingPrices(i) = 0
    
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
            
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1) <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1) <> tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        End If
            
            '3d Increase the tickerIndex.
       
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
        End If
            
        'End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    
    Next i

If stock performance was positive (any outcome above 0%), the resulting cell(s) with positive values were highlighted green. If stock performance was negative (any outcome below 0%), the resulting cells(s) with negative values were highlighted red.

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

The results of the refactored code (in terms of how quckly each Stock Year Analysis was processed) was captured as follows:

<img width="252" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/86274124/126077480-b2a0615c-bb73-4b84-b331-78618ef21f8d.png">
<img width="248" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/86274124/126077482-0166df20-c699-4d90-b11c-9a767728a719.png">


### Summary

  Results garnered from the years analyzed showed that 2017 was a great year for stocks and that 2018 was a bad year for stocks (in relation to the stocks analyzed). A purpose of this challeenge beyond determining stock performance was to refactor VBA code in order to run a smoother process of compiling and outputting the necessary information.
  
  Refactoring code has it's advantages, amongst them being the heightened understanding of the code being formatted. To refactor code means to reformat the code and "clean it up" so that the program runs faster. The better the formatting, the less processes the computer must compute, the faster the output. The speed of output is a big advantage of refactoring as well, as original trials of running Stock Analysis code resulted in code taking 1+ seconds to run. Refactoring of the code led to code taking less than 1 second to run, showing that the refactoring process led to a speedier output. 
  
  A disadvantage of refactoring code is that changing code can lead to errors. One misspelled work, one wrong space, or one wrong reference can stop an entire section of code. Refactoring code led to many headaches and many hands being thrown up in annoyance as it can seem that something that was not broken has now become broken. Changing code can be scary, to potentially mess-up an entire project is a possibility that can induce massive amounts of anxiety in any developer. While original code may not run as fast as refactored code, it is comforting to work with what you know works. 
  
  There were advantages to refactoring this specific instance of VBA code. On top of gaining a greater understanding of VBA code, it led to easier processing of the code. With the refactored code it became possible to run code from one Module for both years without having to create separate scripts. Refactoring this specific VBA code also condensed several different instances of code into one, leading to a smoother transtion to the necessary results. Instead of running several instances of code, only one instance of code need be ran.
  
  The disadvantages to refactoring code in VBA boiled down to many debug errors. The resulting text from a debug error can lead to more questions than answers, along with the highlighted portion of code needing to be addressed not giving a clear insight into what needs to be done. This became such an issue that original code that was leading to errors led to a perfectly runniing output after simply closing and reopening Excel. This caused even greater distress as this can be seen as the equivalent of, "Have you tried turning your computer on and off again?" Beyond these issues and the amount of time needed to refactor code, the disadvantages do not outweigh the advantages of refactoring code.

  The take-away from refactoring code can be as thus: While refactoring code can be daunting, stressful, and can lead to what may seem like more complex coding, the resulting output is a cleaner and better-running form of code/program that culminates in a smoother and more robust result.
