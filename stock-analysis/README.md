
# Stock-Analysis


The overall purpose of this project is to look at different trends on companies in the stock market. 

# Purpose 
In this project, we observe the trends for Steve's data on companies in the stock market. On the Excel spreadsheet, there are thirteen tickers in the stock market. Obviously, there are much more tickers in the stock market. If we want to examine the total daily volume and percentage of each ticker, it might take time to run the code to generate the total daily volume of each ticker. Therefore, we have to refine our code so it will not only generate the desired results, but the code will run faster, take up less memory on the computer, and make the code more readable to the user. 


# Overview of The Project  
In this project we will examine how refactoring code will change the running time of the Visual Basic Script. Refactoring is a great way to make the code more consise, is easy to follow for any programmer, and it takes up less space in memory. It turns out when refactoring the code, the run time is much faster, and it takes less memory to run the code in Visual Basic. We are going to exammime a specific piece of code in the Excel document, run that code and its refactored version, and record the running times it took for the subroutine to run. Then, we are going to look at the advatages and disadvantages of refactoring code. 

# Results
In this project, we are looking at different stock markets for the last year. The drawback is the original code presented for Steve may not work for a very large amount of stocks in the market, which would take more time to execute. In this scenario, we are looking at different stocks and comaring their  total daily volume and return for the years 2017 and 2018. We have computed each stock's total daily volume and percent return by generating the code inside the subroutine, AllStocksAnalysis(). Here is the original code in Visual Basic: 

''' 

    Sub AllStocksAnalysis()

    Dim startTime As Single
    Dim endTime As Single

    yearValue = InputBox("What year would you like to run your analysis on? ")

    startTime = Timer

    'Find number of rows (before both loops)'
    Worksheets("All Stocks Analysis").Activate

    Range("A1").Value = "All Stocks (" + yearValue + ")"

    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Income"
    Cells(3, 3).Value = "Return"

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

    Dim startingPrice As Double
    Dim endingPrice As Double

    Worksheets(yearValue).Activate

    rowCount = Cells(Rows.Count, "A").End(xlUp).Row

    For i = 0 To 11

                ticker = tickers(i)
                totalVolume = 0

                Worksheets(yearValue).Activate


                    For j = 2 To rowCount
                    'Activate Date Worksheet'

                             If Cells(j, 1).Value = ticker Then
                             totalVolume = totalVolume + Cells(j, 8).Value
                            End If

                            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                            startingPrice = Cells(j, 6).Value
                            End If

                            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                            endingPrice = Cells(j, 6).Value
                            End If

                            Next j

                      'Output Results'
                      Worksheets("All Stocks Analysis").Activate
                      Cells(4 + i, 1).Value = ticker
                      Cells(4 + i, 2).Value = totalVolume
                      Cells(4 + i, 3).Value = (endingPrice / startingPrice) - 1
                      Next i

                      dataRowStart = 4
                        dataRowEnd = 15

                        For i = dataRowStart To dataRowEnd

                            If Cells(i, 3).Value > 0 Then

                            Cells(i, 3).Interior.Color = vbGreen

                            ElseIf Cells(i, 3).Value < 0 Then

                            Cells(i, 3).Interior.Color = vbRed

                            Else

                            Cells(i, 3).Interior.Color = xlNone


                        End If

                        Next i
                      'Stop the time to run'
                      endTime = Timer

                      MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

                    'Formatting'
                    'Cells(4, 3).Interior.Color = vbGreen
     End Sub

'''








  

Once we ran this subroutine, we get the following stock outputs for the years 2017 and 2018 

<img width="817" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/104328106/175700224-a1afee31-db3e-46b8-8cf4-b2ba0f2d7e5a.png">

<img width="859" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/104328106/175700239-536728e1-4578-4db6-8a08-ac7fd56a55be.png">

From the subroutine, we have the code running properly and diplaying a considerably fast runtime. In this observation, we would expect the subroutine to run faster when we refactor the original code in AllStockAnalysis(). We have called the refactored version of this subroutine AllStockAnalysisRefactored(). Here is the code for the refactored version of the subroutine AllStocksAnalysis():  


    '''
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

        'Activate data worksheet
        Worksheets(yearValue).Activate

        'Get the number of rows to loop over
        rowCount = Cells(Rows.Count, "A").End(xlUp).Row

        '1a) Create a ticker Index
        tickerIndex = 0

        '1b) Create three output arrays
        Dim tickerVolumes(12) As Long
        Dim tickerStartingPrices(12) As Single
        Dim tickerEndingPrices(12) As Single

        ''2a) Create a for loop to initialize the tickerVolumes to zero.
        For i = 0 To 11

        ticker = tickers(tickerIndex)
        tickerVolumes(tickerIndex) = 0
        Worksheets(yearValue).Activate

        ''2b) Loop over all the rows in the spreadsheet.
        For j = 2 To rowCount

            '3a) Increase volume for current ticker
           If Cells(j, 1).Value = ticker Then

           tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value

           End If

            '3b) Check if the current row is the first row with the selected tickerIndex.
            'If  Then
                 If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

                tickerStartingPrices(tickerIndex) = Cells(j, 6).Value

                End If

            'End If

            '3c) check if the current row is the last row with the selected ticker
             'If the next row’s ticker doesn’t match, increase the tickerIndex.
            'If  Then

                 If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

                tickerEndingPrices(tickerIndex) = Cells(j, 6).Value

                End If

              Next j

                '3d Increase the tickerIndex.
                tickerIndex = tickerIndex + 1


            'End If

        Next i

        '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
        For i = 0 To 11

            Worksheets("All Stocks Analysis").Activate

            Cells(4 + i, 1).Value = tickers(i)
            Cells(4 + i, 2).Value = tickerVolumes(i)
            Cells(4 + i, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i)) - 1

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

'''

When we run this subroutine, we get the following stock outputs and the runtimes on Excel: 

<img width="850" alt="VBA_Challenge_2017_Refactored" src="https://user-images.githubusercontent.com/104328106/175719156-fa270d5b-796e-4fb5-875c-4742ce107621.png">

<img width="877" alt="VBA_Challenge_2018_Refactored" src="https://user-images.githubusercontent.com/104328106/175719169-c3ed4236-c7fa-4498-b15c-b81faafa2668.png">

As we can see, the refactored version of the code does run faster than the original code. 


# Summary 

From the results, it is efficient to refactor the code in order to not only run faster, but take less memory to complete the task. What was interesting was it took more code to refine the subroutine. There are some advantages and disadvantages to refactoring general code. On the one hand, it is easier to understand, and is easier to maintain. On the other hand, it can be time consuming to refactor the code as the lines of code can increase and be cumbersome to deal with. We have to make sure that there are no mistakes and we making sure that it is the right code we need. 

Visual Basic is Microsoft's propriety programming language. Since this is a Microsoft product, it could be challenging to transform programs in Visual Basic to other kinds of operating systems. Visual Basic is easy to learn the syntax. Examining the aforementioned code, the running time was faster, the code was easier to follow, and it took less memory space to run the code. However, the lines of code was larger in the refactored code than the original code. It also took a considerable amount of time to edit and watch out for syntax errors in the code. 

In the original code, the code was not as easy to follow than the refactored code. There are also parts in the code where multiple sheets are activated.  Otherwise, it certainly is readable code, there were minimal lines of code to refine and it did not take as much time as we thought to refine the code. 

The interesting portion of the project is that neither piece of code did not add in a condition where a user inputs anything other than the years described in the project. We had to add in a conditional statement where a user does input something other than the years 2017 and 2018. This is a process in coding called exception handling. What we kept in mind was figuring out ways to catch errors and not make our program crash. 
