VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub mod2()

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

    ' Works for all worksheets
    For Each ws In ThisWorkbook.Worksheets

        ' For current worksheet
        ws.Activate

        ' Last Row
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        '  variables for output
        summaryRow = 2 ' Skipp headers
        volumeTotal = 0
        openPrice = ws.Cells(2, 3).Value ' Open price of the first ticker

        '  headers for output
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Total Volume"
        ws.Cells(1, 11).Value = "Quarterly Change"
        ws.Cells(1, 12).Value = "Percent Change"

        ' Looping through all rows
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

                ' Output  summary table
                ws.Cells(summaryRow, 9).Value = ticker
                ws.Cells(summaryRow, 10).Value = volumeTotal
                ws.Cells(summaryRow, 11).Value = quarterlyChange
                ws.Cells(summaryRow, 12).Value = percentChange

                ' Conditional formatting
                If quarterlyChange >= 0 Then
                    ws.Cells(summaryRow, 11).Interior.Color = RGB(0, 255, 0) ' Green
                Else
                    ws.Cells(summaryRow, 11).Interior.Color = RGB(255, 0, 0) ' Red
                End If

                If percentChange >= 0 Then
                    ws.Cells(summaryRow, 12).Interior.Color = RGB(0, 255, 0) ' Green
                Else
                    ws.Cells(summaryRow, 12).Interior.Color = RGB(255, 0, 0) ' Red
                End If

                '  Next row in the summary table
                summaryRow = summaryRow + 1

                ' Reset variables for next loop
                volumeTotal = 0
                openPrice = ws.Cells(i + 1, 3).Value

            Else
                ' volume for the current ticker
                volumeTotal = volumeTotal + ws.Cells(i, 7).Value
            End If

        Next i

    Next ws

End Sub

