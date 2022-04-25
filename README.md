# Challenge2
Microsoft Excel VBA
#** An Analysis of Refractor VBA Code and Measuring Performance**

## Overview of Project

The overall goal of this project is to compare 2 methods of deriving computed results of ticker (stock symbols), total trading volume and annual returns in an annual report fashion. This comparison was done in the following manner 

1. Analyze daily trading data of 12 exchange listed and over the counter stocks where daily trading dates, trading volume, and open, high and closing prices were displayed for an entire year.

2. In the first method, the original project VBA code (All Stock Analysis) was written such that the stock data was gathered by means of a repeated loop manner going through all of the stock data 12 times, one for each of the stocks in the portfolio.

3. In the second method, (VBA Refactored Code) arrays were created in the VBA code such that the entire stock data could be reviewed in one large loop. 

4. The performance of each method was measured; the conclusion was that the method using arrays was much more efficient. The run time of this method was far superior as measured by the time of execution of each method.

5. This project provided the student the opportunity to gain experience in VBA Programming while understanding the importance of code planning, programming strategy and coding techniques. 


## Refactored Code Segment

The key component of the refactored code are those portions where arrays were created for TickerVolumes, TickerStartingPrices, TickerEndingPrices and TickerIndex was created as a variable, Section 2b shows how the code was looped over all of the rows of the data.

'1a Create a TickerIndex
  tickerindex = 0

'1b Create the 3 Output arrays

  Dim TickerVolumes(0 to 11) as Long
  Dim TickerStartingPrices(0 to 11) as Single
  Dim TickerEndingPrices(0 to 11) as Single

 '2a Create a "For" loop to initialize the TickerVolumes to zero     
  For t = 0 to 11
    TickerVolumes(t)=0   
    TickerStartingPrices(t)=0
    TickerEndingPrices(t)=0
  Next t

 '2b Loop over all of the rows in the spreadsheet                
  For i =2 to rowCount
             
     ' 3a Increase volume for current ticker
      TickerVolumes(TickerIndex) = Cells(i,8).Value+TickerVolumes(TickerIndex)
       
      ' 3b. Check if the current row is the first row with the selected TickerIndex
       ' If then

        if Cells(i,1).Value= Tickers(TickerIndex) and Cells(i-1,1).Value<>Tickers(TickerIndex) Then
           TickerStartingPrices(TickerIndex)= Cells(i,3).Value


## Run Time Comparison by Year: Original Project Code vrs Refactored Code 

The following chart displays the comparison of the run time required by year for each method; the original code provided and the refactored code.


	Original (All Stock Analysis) 
       Project Code	                       VBA Refactored Code

2017	      .55117188	                          .1367188
2018	      .5446875	                          .0703125


   As displayed above the refactored code runs in a much more efficient manner with decreased run times. While this difference may not seem like much in this small program, on a larger scale refactored code
would provide a noticeable difference to users and use of computer resources (ie-servers). 

   The disadvantage of refactored code might be in the time required to restructure the original results. There would be a cost in terms of development time, coding and testing. 

   The above emphasizes the importance of the initial system review and system/coding development to provide efficient performance. 

## Stock Performance/Run Times: Links are provided for screen shots which display the stock portfolio performance by year and run times.


All Stock Analysis Code	
2017	https://user-images.githubusercontent.com/101996041/161670357-2cf8c4ed-324e-494a-8e24-21b7e46323c7.png
2018	https://user-images.githubusercontent.com/101996041/161670376-088ff721-1a7c-4524-a84f-0fb7febd208e.png

VBA Refactored Code	
2017	https://user-images.githubusercontent.com/101996041/161670387-c5e1b4c9-fbc6-463c-af80-143c8087631e.png
2018	https://user-images.githubusercontent.com/101996041/161670403-755e4918-b49c-4ae1-b83d-fff5873bc10f.png



