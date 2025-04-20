VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub AnalyzeQuarterlyStockData()

    Dim ws As Worksheet
    Dim Ticker As String
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim TotalVolume As Double
    Dim QuarterlyChange As Double
    Dim PercentChange As Double
    Dim SummaryRow As Long
    Dim LastRow As Long
    Dim i As Long

    Dim MaxIncrease As Double
    Dim MaxDecrease As Double
    Dim MaxVolume As Double
    Dim IncTicker As String
    Dim DecTicker As String
    Dim VolTicker As String

    For Each ws In Worksheets
        If ws.Name Like "Q*" Then ' Only loop through Q1–Q4 sheets
            ws.Activate

            ' Headers
            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Quarterly Change"
            ws.Range("K1").Value = "Percent Change"
            ws.Range("L1").Value = "Total Volume"

            SummaryRow = 2
            MaxIncrease = -999999
            MaxDecrease = 999999
            MaxVolume = 0
            TotalVolume = 0

            LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            OpenPrice = ws.Cells(2, 3).Value ' column C (open)

            For i = 2 To LastRow

                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    Ticker = ws.Cells(i, 1).Value
                    ClosePrice = ws.Cells(i, 6).Value ' column F (close)
                    TotalVolume = TotalVolume + ws.Cells(i, 7).Value ' column G (volume)

                    QuarterlyChange = ClosePrice - OpenPrice
                    If OpenPrice <> 0 Then
                        PercentChange = QuarterlyChange / OpenPrice
                    Else
                        PercentChange = 0
                    End If

                    ' Summary Output
                    ws.Cells(SummaryRow, 9).Value = Ticker
                    ws.Cells(SummaryRow, 10).Value = QuarterlyChange
                    ws.Cells(SummaryRow, 11).Value = PercentChange
                    ws.Cells(SummaryRow, 12).Value = TotalVolume

                    ' Color Format
                    If QuarterlyChange >= 0 Then
                        ws.Cells(SummaryRow, 10).Interior.Color = RGB(0, 255, 0)
                    Else
                        ws.Cells(SummaryRow, 10).Interior.Color = RGB(255, 0, 0)
                    End If

                    If PercentChange >= 0 Then
                        ws.Cells(SummaryRow, 11).Interior.Color = RGB(0, 255, 0)
                    Else
                        ws.Cells(SummaryRow, 11).Interior.Color = RGB(255, 0, 0)
                    End If

                    ' Max checks
                    If PercentChange > MaxIncrease Then
                        MaxIncrease = PercentChange
                        IncTicker = Ticker
                    End If

                    If PercentChange < MaxDecrease Then
                        MaxDecrease = PercentChange
                        DecTicker = Ticker
                    End If

                    If TotalVolume > MaxVolume Then
                        MaxVolume = TotalVolume
                        VolTicker = Ticker
                    End If

                    SummaryRow = SummaryRow + 1
                    TotalVolume = 0
                    OpenPrice = ws.Cells(i + 1, 3).Value

                Else
                    TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                End If
            Next i

            ' Output Greatest Metrics
            ws.Range("O1").Value = "Metric"
            ws.Range("P1").Value = "Ticker"
            ws.Range("Q1").Value = "Value"

            ws.Range("O2").Value = "Greatest % Increase"
            ws.Range("P2").Value = IncTicker
            ws.Range("Q2").Value = FormatPercent(MaxIncrease, 2)

            ws.Range("O3").Value = "Greatest % Decrease"
            ws.Range("P3").Value = DecTicker
            ws.Range("Q3").Value = FormatPercent(MaxDecrease, 2)

            ws.Range("O4").Value = "Greatest Volume"
            ws.Range("P4").Value = VolTicker
            ws.Range("Q4").Value = MaxVolume
        End If
    Next ws

    MsgBox "? Full quarterly stock analysis complete!"

End Sub

