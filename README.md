# Stock Analysis in Excel Using VBA

## Overview of Project
Steve is impressed with the stock analysis worksheet I created to analyze the green stocks his parents were interested in. He tasked me with creating a sheet and VBA code that could analyze the whole stock market quickly and efficiently. 

### Purpose
The purpose of this project is to take my existing VBA code which was run on a handful of stocks and make it be able to handle thousands of stocks as quickly as possible.

## Analysis and Results

### Stock Analysis and Results
Looking at the results of the code for the different years show very different results in terms of returns. After tremendous  gains in 2017, 2018 was a very bad year for most green stocks. There is an exception for two companies, tickers $ENPH and $RUN which had positive returns in 2018. In terms of volume from 2017 to 2018 volume increase on most stocks which. Tables showing the total volume and % return for 2017 and 2018 can be seen below.

Table 1: 2017 Stock Analysis

![VBA_Challenge_2017](https://user-images.githubusercontent.com/57120024/156947438-3bccb166-7daf-4db0-9174-25ae0762acc9.png)

Table 2: 2018 Stock Analysis

![VBA_Challenge_2018](https://user-images.githubusercontent.com/57120024/156947449-e9114d12-83c1-43ee-9ba2-cb3605c5a412.png)

### VBA Code Results
When comparing the new refactored code to my original code, the refactored code will now work with data from any year of the 12 green stocks we looked at before. When it comes to how fast the code is executed, there is a great improvment in time to execute. The biggest improvement in speed came from not using a nested for loop and instead defining a tickerIndex variable and increasing it at the end of the loop looking for total Volume, Starting Prices, and Ending Prices and using a seperate for loop to set all elements in the tickerVolumes array to 0: 

    Dim tickerIndex As Single
      tickerIndex = 0 
     
    For i = 0 To 11
      tickerVolumes(i) = 0
    Next i  
    
    For j = 2 To RowCount
            
        If Cells(j, 1).Value = tickers(tickerIndex) Then
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
        End If

        If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j - 1, 1).Value <> tickers(tickerIndex) Then
            startingPrice(tickerIndex) = Cells(j, 6).Value
        End If

        If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
            endPrice(tickerIndex) = Cells(j, 6).Value
            tickerIndex = tickerIndex + 1
        End If
    Next j
     
      
    

![2017 original code time](https://user-images.githubusercontent.com/57120024/156947684-71ca95c1-bdd0-4855-8570-88494dd29295.PNG)

Figure 1: 2017 original code time

![2017_refactored2_time](https://user-images.githubusercontent.com/57120024/157288436-42fd35ee-4948-4d06-8d18-663575c327b3.PNG)

Figure 1: 2017 refactored code time

![2018 original code time](https://user-images.githubusercontent.com/57120024/156947690-fb8706d7-8f30-4a3c-bfce-a9e17f7522ba.PNG)

Figure 1: 2018 original code time

![2018_refactored2_time](https://user-images.githubusercontent.com/57120024/157288474-38f9eeec-ee28-4d2f-829b-4be93303a946.PNG)

Figure 1: 2018 refactored code time

## Summary

1) What are the advantages or disadvantages of refactoring code?

  There and many advantages of refactoring  your code. The main advantage  is to make it more useful and less specific to the original data you were working with. Another advatage is making the code execute more efficiently and quicker, this is important when working with large data sets. 
  
  While there are advantaged there are also disadvantages. The main disadvantages is it could complicate your code. When refactoring the code you run into bugs which is time consuming and frustrating after creating working code. 
  
2) How do these pros and cons apply to refactoring the original VBA script?

  When it came to my original VBA script, refactoring did greatly improve/lower the time to execute. Also now the script can run for more years of stocks data with the inclusion of the text input, for example we can look at 2019 data even if it is not complete for the year. The code for the input box is shown below:
  
      "yearValue = InputBox("What year would you like to run the analysis on?")"
      
 The script is still limited to the 12 green stocks due to the tickers array being manually populated with the stock tickers. A way to make the code even more is to have the code read through the column of stock tickers and fill an array with unique stock tickers. This would increase the time to execute and add a bit of complexity to the code but would greatly improve the usefulness of the code. 
      
