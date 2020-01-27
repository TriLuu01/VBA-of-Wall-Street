Attribute VB_Name = "Module2"
Sub AllStocksAnalysis()
    yearValue = InputBox("What year would you like to run the analysis on?")

    Worksheets("AllStocksAnalysis").Activate

    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    'Creating 4 arrays
    Dim tickers(12) As String
    Dim startingPrice(12) As Double
    Dim endingPrice(12) As Double
    Dim TotalVolume(12) As Double
    'Creating index variable
    Dim tickerIndex As Integer
    'Set the value equal to zero before the loops
    tickerIndex = 0
    
    Worksheets(yearValue).Activate

    'get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    For tickerIndex = 0 To 11

        TtVolume = 0

        Worksheets(yearValue).Activate

        'loop over all the rows
        For j = 2 To RowCount
            'getting data for starting price array
            If Cells(j - 1, 1).Value <> Cells(j, 1).Value Then
                tickers(tickerIndex) = Cells(j, 1).Value
                startingPrice(tickerIndex) = Cells(j, 6).Value

            End If
            
            If Cells(j, 1).Value = tickers(tickerIndex) Then
                TotalVolume(tickerIndex) = TotalVolume(tickerIndex) + Cells(j, 8).Value
               
            End If
            
            'getting data for ending price array
            If Cells(j + 1, 1).Value <> Cells(j, 1).Value Then
                
                endingPrice(tickerIndex) = Cells(j, 6).Value
                tickerIndex = tickerIndex + 1
            End If

        Next j

    Next tickerIndex
    
    'Output the all the arrays' data
    Worksheets("AllStocksAnalysis").Activate
    For i = 0 To 11
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = TotalVolume(i)
        Cells(4 + i, 3).Value = endingPrice(i) / startingPrice(i) - 1
    Next i
    'Formatting
    Worksheets("AllStocksAnalysis").Activate
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

End Sub
