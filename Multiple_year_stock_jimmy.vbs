Attribute VB_Name = "Module1"
Sub Practice3()
    Dim ws As Worksheet
    Dim ticker As String
    Dim Summary_Table_Row As Integer
    Dim Lastrow As Long
    Dim i As Long
    Dim startDate As Date
    Dim endDate As Date
    Dim startPrice As Double
    Dim endPrice As Double
    Dim quarterlyChange As Double
    Dim percentageChange As Double
    Dim TotalStockVolume As Double
    Dim rng As Range
    Dim greatestIncreaseTicker As String
    Dim greatestIncreasevalue As Double
    Dim greatestDecreaseTicker As String
    Dim greatestDecreaseValue As Double
    Dim greatestVolumeTicker As String
    Dim greatestIncreaseVolume As Double
    
    For Each ws In ThisWorkbook.Worksheets
        Summary_Table_Row = 2
        TotalStockVolume = 0
        greatestIncreasevalue = 0
        greatestIncreaseTicker = " "
        greatestDecreaseValue = 0
        greatestDecreaseTicker = " "
        greatestVolumeTicker = " "
        greatestIncreaseVolume = 0
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percentage Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"


        Lastrow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

        For i = 2 To Lastrow
            ticker = ws.Cells(i, 1).Value
        
            If i = 2 Or ticker <> ws.Cells(i - 1, 1).Value Then
                startDate = ws.Cells(i, 2).Value
                startPrice = ws.Cells(i, 3).Value
            End If
        
            endDate = ws.Cells(i, 2).Value
            endPrice = ws.Cells(i, 6).Value
            TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value

            If i = Lastrow Or ticker <> ws.Cells(i + 1, 1).Value Then
                quarterlyChange = endPrice - startPrice
                percentageChange = ((endPrice - startPrice) / startPrice) * 100
                percentageChange = Round(percentageChange, 2)
                ws.Cells(Summary_Table_Row, 9).Value = ticker
                ws.Cells(Summary_Table_Row, 10).Value = quarterlyChange
                ws.Cells(Summary_Table_Row, 11).Value = percentageChange & "%"
                ws.Cells(Summary_Table_Row, 12).Value = TotalStockVolume
                Summary_Table_Row = Summary_Table_Row + 1
                TotalStockVolume = 0
           
            If percentageChange > greatestIncreasevalue Then
                greatestIncreasevalue = percentageChange
                greatestIncreaseTicker = ticker

            End If

            If percentageChange < greatestDecreaseValue Then
                greatestDecreaseValue = percentageChange
                greatestDecreaseTicker = ticker

            End If
         End If
        
            If TotalStockVolume > greatestIncreaseVolume Then
                greatestIncreaseVolume = TotalStockVolume
                greatestVolumeTicker = ticker

            End If
        Next i

        ws.Cells(2, 16).Value = greatestIncreaseTicker
        ws.Cells(2, 17).Value = greatestIncreasevalue & "%"
        ws.Cells(3, 16).Value = greatestDecreaseTicker
        ws.Cells(3, 17).Value = greatestDecreaseValue & "%"
        ws.Cells(4, 16).Value = greatestVolumeTicker
        ws.Cells(4, 17).Value = greatestIncreaseVolume
        
        Set rng = ws.Range(ws.Cells(2, 10), ws.Cells(Lastrow, 10))
        
        For Each Cell In rng
        
            If Cell.Value > 0 Then
                Cell.Interior.Color = RGB(0, 255, 0)
            ElseIf Cell.Value < 0 Then
                    Cell.Interior.Color = RGB(255, 0, 0)
            ElseIf Cell.Value = 0 Then
                    Cell.Interior.Color = RGB(255, 255, 255)
             End If
             
        Next Cell
    Next ws
End Sub
