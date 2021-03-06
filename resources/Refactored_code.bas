Attribute VB_Name = "Module12"
Sub Btn_Run_refactored_stockanalysis_Click()
  
    'Declare the variables for stock opening time and closing time resp
    Dim startTime As Single
    Dim endTime  As Single

    'Accept user input for year to run the analysis on
    yearValue = InputBox("What year would you like to run the analysis on?")

    'Start timer to track execution run time
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
    rowcount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    
    tickerIndex = 0
    
    '1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12), tickerEndingPrices(12) As Single
    
      
    '2a) Create a for loop to initialize the tickerVolumes to zero.

    For tickerIndex = 0 To 11
        tickerVolumes(tickerIndex) = 0
        tickerStartingPrices(tickerIndex) = 0
        tickerEndingPrices(tickerIndex) = 0
    Next tickerIndex
    
    tickerIndex = 0
        
    '2b) Loop over all the rows in the spreadsheet.
    For i = 2 To rowcount
    
        '3a) Increase volume for current ticker
                      
        ticker = tickers(tickerIndex)
        
        If ticker = Cells(i, 1).Value Then
        
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
            '3b) Check if the current row is the first row with the selected tickerIndex.
        
             'If Cells(i, 1).Value = ticker And Cells(i - 1, 1).Value <> ticker Then
              If Cells(i - 1, 1).Value <> ticker Then
               
               tickerStartingPrices(tickerIndex) = Cells(i, 6).Value

              End If
                    
              
            '3c) check if the current row is the last row with the selected ticker
        
              'If Cells(i, 1).Value = ticker And Cells(i + 1, 1).Value <> ticker Then
              If Cells(i + 1, 1).Value <> ticker Then

               tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
               
               '3d) If the next row's ticker doesn't match, increase the tickerIndex.
               
               tickerIndex = tickerIndex + 1

              End If
            
        End If
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    
        
        Worksheets("All Stocks Analysis").Activate
        
        For i = 0 To 11
               
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
     
        Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    'Make the Font Bold
    Range("A3:C3").Font.FontStyle = "Bold"
    'Give an outline on bottom edge of the cell range
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    'Assign number format to TotalVolume column
    Range("B4:B15").NumberFormat = "#,##0"
    'Assign percentage format to Return column
    Range("C4:C15").NumberFormat = "0.0%"
    'Shrink the column to fit the data
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    'Assign conditional formatting to show -ve returns as Red and +ve returns as green
    
    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    'Stop timer and Display execution time
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub


