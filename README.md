# Module 2 - VBA Challeng

## Overview of Project 
### Purpose 
The purpose of this analysis was to take the workbook that was prepared for Steve and refactor it, giving it more capabilities and making it more efficient. Both solutions resulted in giving Steve the total volume and return of each of the ticker types.
## Results 
### Analysis 
The code underwent several changes in order to have the capability to run through more stocks at a faster pace. The first is that a ticker index was created. The purpose of the ticker index, along with the code behind it, was to make it so the code only had to run through the data set once and check for all the ticker types. Previously, the code run through the list 12 times for each individual ticker. With this, the code became more efficient, and it did not require a nested "For" loop to get the same results. See Final code and time results below.  

  
### All Stocks Analysis Refactored
    
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
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    'Setting ticker index to 0 allows for the index to start fresh and add value as it cycles through to identify the correct ticker
    
    tickerIndex = 0

    '1b) Create three output arrays
    'Code creates output arrays and defines variable types
    
    Dim tickerVolumes(11) As Long
    Dim tickerStartingPrices(11) As Single
    Dim tickerEndingPrices(11) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    'Initilizes the volume, startingprices and endprices to 0 when the loop starts on a new ticker type
    
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
        'COde checks the row above to see if a new ticker type has been started so it knows the starting price
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1) = tickers(tickerIndex) Then
             
             tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            
        End If

        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'Code checks the row above to see if a ticker type has ended so it knows the end prices
            
         If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
         
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
         End If

            '3d Increase the tickerIndex.
            'Increases ticker index so we can cycle through the tickers
            
         If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1) = tickers(tickerIndex) Then
        
            tickerIndex = tickerIndex + 1
        
        End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    'Outputs results
    
    For i = 0 To 11
        
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
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

### Refractored Time Results
<img width="239" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/111031608/186537687-ecaa5b2d-512f-451b-8f9e-dab8c74008f6.PNG">
<img width="241" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/111031608/186537692-a6105119-6456-4b52-96e5-508f69ff3acc.PNG">

## Summary 
### Advantages of Refractoring Code
There are several advantages to refactoring code. Refactoring can give a program greater capability with less power and time. Another benefit to refactoring code is the potential for it to be written clearer and more concise. This is helpful because it makes the code easier to read and detect errors especially for outside viewers. In the VBA code created for Steve, all these benefits apply. We removed a "For" loop, gave the code more capabilities, and saved the amount of time it took to run the code. 
### Disadvantage of Refactoring Code
Refactoring code is a double-edged sword. Since code can be written many ways, it can be refactored to be more efficient, or it can have the opposite effect. Sometimes to refactor code, more bugs and issues are introduced. If not done correctly, refactoring can have the opposite effect intended. With Steve’s VBA Analysis, the code had several bugs that needed to be worked through before finally becoming a working code that had an original effect. This process took a lot of time and energy that would have been wasted if the code was unable to run. 




