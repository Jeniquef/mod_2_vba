VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub StockAnalysis()

    ' Variables
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim volumeTotal As Double
    Dim summaryRow As Integer
    Dim i As Long
    Dim quarterlyChange As Double
    Dim percentChange As Double
    
    ' Variables to track greatest values
    Dim maxIncrease As Double
    Dim maxIncreaseTicker As String
    Dim maxDecrease As Double
    Dim maxDecreaseTicker As String
    Dim maxVolume As Double
    Dim maxVolumeTicker As String

    ' Initialize tracking variables
    maxIncrease = -100000
    maxDecrease = 100000
    maxVolume = 0

    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets

        '  current worksheet
        ws.Activate

        ' Last Row
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        ' Initialize summary table
        summaryRow = 2
        volumeTotal = 0
        openPrice = ws.Cells(2, 3).Value ' First open price

        ' Headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        ' Loop through rows
        For i = 2 To lastRow
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                ' Get ticker
                ticker = ws.Cells(i, 1).Value

                ' Get closing price
                closePrice = ws.Cells(i, 6).Value

                ' Calculations
                volumeTotal = volumeTotal + ws.Cells(i, 7).Value
                quarterlyChange = closePrice - openPrice
                If openPrice <> 0 Then
                    percentChange = (quarterlyChange / openPrice) * 100
                Else
                    percentChange = 0
                End If

                ' Output summary
                ws.Cells(summaryRow, 9).Value = ticker
                ws.Cells(summaryRow, 10).Value = quarterlyChange
                ws.Cells(summaryRow, 11).Value = percentChange
                ws.Cells(summaryRow, 12).Value = volumeTotal

                ' Apply Conditional Formatting
                If quarterlyChange >= 0 Then
                    ws.Cells(summaryRow, 10).Interior.Color = RGB(0, 255, 0) ' Green
                Else
                    ws.Cells(summaryRow, 10).Interior.Color = RGB(255, 0, 0) ' Red
                End If

                ' Find Greatest Increase
                If percentChange > maxIncrease Then
                    maxIncrease = percentChange
                    maxIncreaseTicker = ticker
                End If

                ' Find Greatest Decrease
                If percentChange < maxDecrease Then
                    maxDecrease = percentChange
                    maxDecreaseTicker = ticker
                End If

                ' Find Greatest Volume
                If volumeTotal > maxVolume Then
                    maxVolume = volumeTotal
                    maxVolumeTicker = ticker
                End If

                ' Move to next summary row
                summaryRow = summaryRow + 1

                ' Reset for next ticker
                volumeTotal = 0
                openPrice = ws.Cells(i + 1, 3).Value

            Else
                '  volume total
                volumeTotal = volumeTotal + ws.Cells(i, 7).Value
            End If

        Next i

        '  Increase, Decrease, and Volume Output
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"

        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(2, 15).Value = maxIncreaseTicker
        ws.Cells(2, 16).Value = Format(maxIncrease, "0.00%")

        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(3, 15).Value = maxDecreaseTicker
        ws.Cells(3, 16).Value = Format(maxDecrease, "0.00%")

        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(4, 15).Value = maxVolumeTicker
        ws.Cells(4, 16).Value = Format(maxVolume, "0.00E+00")

        ' Apply formatting
        ws.Columns("I:L").AutoFit
        ws.Columns("O:Q").AutoFit

    Next ws

End Sub

