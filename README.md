# stock-analysis

Module 2 VBA Chanllenge with Refactor code of stock-analysis.

New assignment consists of one technical deliverable and a written report
Deliverable 1: Refactor VBA code and measure performance
		'This deliverable will include an updated workbook and a folder with PNGs of the pop-ups with script run time
Deliverable 2: A written analysis of your results (README.md)

## Overview of Project Stock Analysis

Performning a refractor of VB code created for The Stock Analysis project, this project will also measure performance.

The Stock Analysis project began with assisting Steve and his parents request to analyze performance if their investment of a green stock DAQO New Energy Corp. DAQO's ticker symbol is "DQ". 

In order to analyze stock performance of "DQ" stock, data was collected for eleven additional green stocks for the years of 2017 and 2018.  

Both Steve and his parents requested additional analysis to include all twelve green stocks options to find out returns.  

Visual Basic for Applications or VBA was implmented in the analysis of the stock data and is considered a good tool for the financial industry due to it's ease of automating tasks while reducing both errors and time required to run analysis. 


## Purpose
Performning a refractor of VB code created for The Stock Analysis project that will also measure performance between original macro created against refractor macro. To view VB code open VBA_Challenge.xlsm and use Developer module.  
Original macro  named AllStocksAnalysis
Refractor macro named AllStocksAnalysisRefactored
Additional macros included are DQ, ClearWorksheet, oldStockAnslysis, rowsbloops and skilldrill - all part of Module 2 work. 

## DQ Analysis 

Steve requested the total daily volume and yearly return for each stock.  The yearly return is the percentage difference in price from the beginning of the year to the end of the year.
 
Since his parents had invested in DQ, the stock was the first to be analysis.  The analysis of stock for DQ found that the stock for DAQO New Energy Corp did have a great return in 2017 with 199.45% but 2018 not have good returns for 2018. The end of year return found a decrease of price of -62% 
![DQ_20172018](DQ_20172018.png)

Below is the original DQ Analysis code.
Code for DQ Analysis
![DQAnalysis.png](DQAnalysis.png)

## All Stock Analysis 
Now to modfiy code to show all stocks. As per Steve and his parents request, I will analyze multiple stocks options to find out good return for them
We define new sub as All stocks analysis for further coding important part that shows loop through tickers

'5) Loop through rows in the data.

Sheets(yearValue).Activate
    For j = 2 To RowCount
   
'5a) Get the total volume for the current ticker.
    If Cells(j, 1).Value = ticker Then
    
        totalVolume = totalVolume + Cells(j, 8).Value
        
     End If
        
'5b) Get the starting price for the current ticker.

    If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
    
        startingPrice = Cells(j, 6).Value
        
    End If
    
    
'5c)Find the ending price for the current ticker.
    
    If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
    
        endingPrice = Cells(j, 6).Value
        
    End If
	

## Results
The result analysis of all stocks by calulating the total daily volume and yearly return, a couple of good green stocks investment would be stock ENPH with 81.9% and RUN with 84% returns.

![2018.png](2018.png)


## VBA Code Measure Performance
Refactoring code we do analysis to understand what to invest in. -Refactoring is a key part of the coding process. Below is the refactor code:

Sub AllStocksAnalysisRefactored()
==============================================================================
	Dim startTime As Single
	Dim endTime  As Single

	yearValue = InputBox("What year would you like to run the analysis on?")

	startTime = Timer

	'Format the output sheet on All Stocks Analysis worksheet
	Worksheets("All Stocks Analysis").Activate

	Range("A1").Value = "All Stocks (" + yearValue + ")"

	'Create a header row
	cells(3, 1).Value = "Ticker"
	cells(3, 2).Value = "Total Daily Volume"
	cells(3, 3).Value = "Return"

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
	RowCount = cells(Rows.Count, "A").End(xlUp).Row

	'1a) Create a ticker Index
   
   	tickerIndex = 0


	'1b) Create three output arrays

   	Dim tickerVolumes(12) As Long
   	Dim tickerStartingPrices(12) As Single
   	Dim tickerEndingPrices(12) As Single


	'2a) Create a for loop to initialize the tickerVolumes to zero.

   	For i = 0 To 11
   
    	tickerVolumes(i) = 0
    
   	Next i
    
    	'Activate data worksheet
   	Worksheets(yearValue).Activate
    
	'2b) Loop over all the rows in the spreadsheet.

    	For j = 2 To RowCount

    	'3a) Increase volume for current ticker
    
     	For tickerIndex = 0 To 11
     
      If cells(j, 1).Value = tickers(tickerIndex) Then
      
      tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + cells(j, 8).Value
      
      End If
    
    	'3b) Check if the current row is the first row with the selected tickerIndex.
    	'If  Then
        
        If cells(j - 1, 1).Value <> tickers(tickerIndex) And cells(j, 1).Value = tickers(tickerIndex) Then
        
        tickerStartingPrices(tickerIndex) = cells(j, 6).Value
        
        
       End If
    
    	'3c) check if the current row is the last row with the selected ticker
     	'If the next row's ticker doesn't match, increase the tickerIndex.
    	'If  Then
        
        If cells(j + 1, 1).Value <> tickers(tickerIndex) And cells(j, 1).Value = tickers(tickerIndex) Then
        
        tickerEndingPrices(tickerIndex) = cells(j, 6).Value
        
        End If
   
    	'3d Increase the tickerIndex.
    
     	Next tickerIndex
         
    	Next j

	'4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.

     Worksheets("All Stocks Analysis").Activate
    	For i = 0 To 11
    
     cells(4 + i, 1).Value = tickers(i)
     cells(4 + i, 2).Value = tickerVolumes(i)
     cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
     
	 Next i
 
  	Worksheets("All Stocks Analysis").Activate

	endTime = Timer
	MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

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
    
    	If cells(i, 3) > 0 Then
        
     	cells(i, 3).Interior.Color = vbGreen
        
    	Else
    
        cells(i, 3).Interior.Color = vbRed
        
    	End If
    
	Next i

	End Sub
		

==============================================================================
The code was further modified so Steve could run the analysis for either 2017 or 2018 on all stocks. Results include comparison of refactored code to original code. 

Below shows 2017 comparison: 

2017 Refactored data measurement.
![Refactored 2017](VBA_Challenge_2017.png)

2017 Original data measurement.
![Original 2017](All_Stocks_2017.png)


Below shows 2018 comparison: 

2018 data mesurement
![Redfactored 2018](VBA_Challenge_2018.png)

2018 Original data measurement.
![Original 2018](All_Stocks_2017.png)

We can finally conclude that refactoring code screen running time is less then the original script for both the years and the refactoring is more easier to understand and read.


#Summary 
The VB code created can be modified so if in the future both Steve and his parents would like to run analysis of future stock investments it can be done.  The macro for running multiple stocks and different years is very flexible. 
Included in the code is a program that can quickly loop throug all the tickers. Then additional format of tables was included to make stock reasults eaiser to read and used the color of red and green.  

What are the advantages or disadvantages of refactoring code?
The advantage of using and refactoring code is I am able to produce analysis and run the  code  faster and utilize less resources.  Also refactor code is easier to read and follow.
The only disadvantage is to refactor or write code is time consuming and very complex in VB.

'How do these pros and cons apply to refactoring the original VBA script?
The refactoring of code was complex structure and had to google and use youtube to fully complete the assignment.

