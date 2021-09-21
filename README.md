# Stock-Analysis

## Overview of Project
In this analysis I used my prowess in VBA excel to analyze an entire dataset that was filled with a list of 12 stocks from the year 2017 & 2018. My goal was to find what stocks had the highest trading volume which indicated interest. Next I wanted to analyze out of the 12 stocks which had the highest rate of return over the year. After coding in VBA to automate this for me I needed to optimize the code to iprove te speed and performance. 

### Results
![2017](https://user-images.githubusercontent.com/83085800/134173339-c9387b70-5080-48d5-a92c-8d7c0d3cfa51.png)
![VBA_Challenge_2017](https://user-images.githubusercontent.com/83085800/134173342-3b5cecfe-85cb-4a24-a62d-72f714df3ac9.png)
I was able to get the code perfomance faster extremely when I refactored the code.
'1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    ' If the next row’s ticker doesn’t match, increase the tickerIndex.
'    For i = 0 To 11
'        tickerVolumes(i) = 0
'    Next i
   
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
        'If  Then
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
         End If

        '3d Increase the tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
        End If
    
    Next i
   ![2018](https://user-images.githubusercontent.com/83085800/134176809-cfb22805-8a13-4c8b-8d4e-833334203ee8.png)
   ![VBA_Challenge_2018](https://user-images.githubusercontent.com/83085800/134176836-45dfdccb-da8e-4233-9420-c863cd570235.png)
   
   #### Summary
   To begin the summary I want to discuss the disadvantages to refactoring code is mainly time consuming. I had to read through multiple sources to find the best way to go about it. Anything programming is time consuming so it comes with the nature. So you have to defintely love it, which i do. The advantages is the code becomes easier to undersatnd and the code is faster. 


