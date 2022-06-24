# Stock Analysis with VBA

## Overview of Project
    The project focuses on building the skills related to the use and development of VBA in Microsoft Office Applications, specifically Excel. As a part of the project there will tasks to expose and develop skills related to the Macros and different elements that can be manipulated by these tasks. 

### Purpose
    The purpose of the project is to gain basics of code development prior to going into python. Exposure to processes that allow for formatting, setup and design of outputs, and loops to extract and manipulate the data into the different tasks required by the challenge.

## Analysis and Challenges
    The data represented 12 stocks for two years (2017 & 2018) and their daily volumes and prices. The task is to evaluate the stocks for price increase and process the data effectively. I worked through the steps in the activity to test and learn different aspects.

    Sub AllStocksAnalysis()

        '1) Format the output sheet on All Stocks Analysis worksheet
    
        Dim startTime As Single
    
        Dim endTime  As Single

            startTime = Timer
    
        Worksheets("All Stocks Analysis").Activate
    
        yearValue = InputBox("What year would you like to run the analysis on?")
    
        Range("A1").Value = "All Stocks (" + yearValue + ")"
    
        'Create a header row
    
        Cells(3, 1).Value = "Ticker"
    
        Cells(3, 2).Value = "Total Daily Volume"
    
        Cells(3, 3).Value = "Return"
    

        '2) Initialize array of all tickers
    
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
        
            totalVolume = 0
       
       
            '5) loop through rows in the data
    
            Worksheets(yearValue).Activate
        
            For j = 2 To RowCount
       
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

    End Sub

    Sub formatAllStocksAnalysisTable()

        'Formatting
    
        Worksheets("All Stocks Analysis").Activate
    
        Range("A3:C3").Font.Bold = True
    
        Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    
        Range("B4:B15").NumberFormat = "#,##0"
    
        Range("C4:C15").NumberFormat = "0.0%"
    
        Columns("B").AutoFit
    
        dataRowStart = 4
    
        dataRowEnd = 15
    
        For i = dataRowStart To dataRowEnd

            If Cells(i, 3) > 0 Then

                'Color the cell green

                Cells(i, 3).Interior.Color = vbGreen

            ElseIf Cells(i, 3) < 0 Then

                'Color the cell red

                Cells(i, 3).Interior.Color = vbRed

            Else

                'Clear the cell color
                
                Cells(i, 3).Interior.Color = xlNone

            End If

        Next i
    
        endTime = Timer
    
        MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
    
    End Sub

## Results

    Refactoring the code from the original state above resulted in significant improvement in performance.

    By using the best practices to better process data this opens up the opportunities to use these practices to manipulate large amounts of data in a short period of time.

    2017 Performance Improvement
        -Original: 60906.45 seconds
        -Refactor: 0.2460938 seconds

    [Original 2017]https://github.com/plymburner/resources/blob/main/2017%20Original.png
    [VBA_Challenge_2017]https://github.com/plymburner/resources/blob/main/VBA_Challenge_2017.png
    
    2018 Performance Improvement
        -Original: 61157.57 seconds
        -Refactor: 0.2109375 seconds

    [2018 Original]https://github.com/plymburner/resources/blob/main/2018%20Original.png
    [VBA_Challenge_2018]https://github.com/plymburner/resources/blob/main/VBA_Challenge_2018.png    