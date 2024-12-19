Attribute VB_Name = "Module1"
Sub Challenge2()
    ' Define all variables
    Dim ws As Worksheet
    Dim i As Long
    Dim ticker As String
    Dim quarterlyChange As Double
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim percentageChange As Double
    Dim totalVolume As Double
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim highestVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim highestVolumeTicker As String
    Dim outputRow As Long

    ' Initialize variables
    greatestIncrease = 0
    greatestDecrease = 0
    highestVolume = 0
    totalVolume = 0
    outputRow = 2 ' Start output from the second row

    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Set up headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percentage Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 15).Value = "Metric"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"

        ' Calculate last row of data
        Dim lastRow As Long
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

        ' Reset tracking variables for each worksheet
        totalVolume = 0
        greatestIncrease = 0
        greatestDecrease = 0
        highestVolume = 0


        ' Reset openingPrice for first row
        openingPrice = ws.Cells(2, 3).Value

        ' Loop through each row
        For i = 2 To lastRow
            ' Get the ticker symbol
            

            ' Check if we are at the last row of the current ticker
            If i = lastRow Or ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
            
                ticker = ws.Cells(i, 1).Value
            
            ' Set opening and closing prices
            
               
                closingPrice = ws.Cells(i, 6).Value

                ' Calculate quarterly change
                quarterlyChange = closingPrice - openingPrice

                ' Calculate percentage change
                If openingPrice <> 0 Then
                percentageChange = (quarterlyChange / openingPrice)
                Else
                    percentageChange = 0
                End If

                ' Calculate total volume
                totalVolume = totalVolume + ws.Cells(i, 7).Value

                ' Record results in the summary table
                ws.Cells(outputRow, 9).Value = ticker
                ws.Cells(outputRow, 10).Value = quarterlyChange
                ws.Cells(outputRow, 10).NumberFormat = "0.00"
                ws.Cells(outputRow, 11).Value = percentageChange
                ws.Cells(outputRow, 11).NumberFormat = "0.00%"
                ws.Cells(outputRow, 12).Value = totalVolume

                ' Apply conditional formatting for quarterly change
                Select Case quarterlyChange
                    Case Is > 0
                        ws.Cells(outputRow, 10).Interior.ColorIndex = 4 ' Green
                    Case Is < 0
                        ws.Cells(outputRow, 10).Interior.ColorIndex = 3 ' Red
                    Case Else
                        ws.Cells(outputRow, 10).Interior.ColorIndex = 0 ' No color
                End Select

                ' Update greatest increase, decrease, and volume
                If percentageChange > greatestIncrease Then
                    greatestIncrease = percentageChange
                    greatestIncreaseTicker = ticker
                End If

                If percentageChange < greatestDecrease Then
                    greatestDecrease = percentageChange
                    greatestDecreaseTicker = ticker
                End If

                If totalVolume > highestVolume Then
                    highestVolume = totalVolume
                    highestVolumeTicker = ticker
                End If

                ' Increment the output row
                outputRow = outputRow + 1

                ' Reset total volume for the next ticker
                If i < lastRow Then
                    openingPrice = ws.Cells(i + 1, 3).Value
                quarterlyChange = 0
                totalVolume = 0
                End If
                
                
                
                
            Else
                ' Accumulate volume for the current ticker
                totalVolume = totalVolume + ws.Cells(i, 7).Value
            End If
        Next i

        ' Record greatest metrics in summary
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(2, 16).Value = greatestIncreaseTicker
        ws.Cells(2, 17).Value = greatestIncrease
        ws.Cells(2, 17).NumberFormat = "0.00%"

        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Value = greatestDecreaseTicker
        ws.Cells(3, 17).Value = greatestDecrease
        ws.Cells(3, 17).NumberFormat = "0.00%"

        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(4, 16).Value = highestVolumeTicker
        ws.Cells(4, 17).Value = highestVolume
    Next ws

End Sub

