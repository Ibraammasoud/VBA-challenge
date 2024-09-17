Attribute VB_Name = "Module1"
    Sub Loooooop()
    
    Dim ws As Worksheet
    Dim Ticker_Name As String
    Dim Summary_row As Integer
    Dim total_vol As Double
    Dim open_val As Double
    Dim close_val As Double
    Dim qtr_chg As Double
    Dim per_chg As Double
    Dim max_increase As Double
    Dim max_decrease As Double
    Dim max_volume As Double
    Dim max_increase_ticker As String
    Dim max_decrease_ticker As String
    Dim max_volume_ticker As String
    total_vol = 0
    Summary_row = 1
    For Each ws In Worksheets
        max_increase = 0
        max_decrease = 0
        max_volume = 0
        total_vol = 0
        Summary_row = 1
        If ws.Range("J1").Value = "" Then
            ws.Range("J1").Value = "Ticker"
        End If
        If ws.Range("K1").Value = "" Then
            ws.Range("K1").Value = "Quarterly Change"
        End If
        If ws.Range("L1").Value = "" Then
            ws.Range("L1").Value = "Percent Change"
        End If
        If ws.Range("M1").Value = "" Then
            ws.Range("M1").Value = "Total Stock Volume"
        End If
        For i = 2 To ws.Cells(Rows.Count, "A").End(xlUp).Row
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker_Name = ws.Cells(i, 1).Value
                total_vol = total_vol + ws.Cells(i, 7).Value
                close_val = ws.Cells(i, 6).Value
                Summary_row = Summary_row + 1
                ws.Cells(Summary_row, 10).Value = Ticker_Name
                qtr_chg = close_val - open_val
                ws.Cells(Summary_row, 11).Value = qtr_chg
                If qtr_chg > 0 Then
                    ws.Cells(Summary_row, 11).Interior.ColorIndex = 4
                ElseIf qtr_chg < 0 Then
                    ws.Cells(Summary_row, 11).Interior.ColorIndex = 3
                End If
                If open_val <> 0 Then
                    per_chg = qtr_chg / open_val
                Else
                    per_chg = 0
                End If
                ws.Cells(Summary_row, 12).Value = per_chg
                ws.Cells(Summary_row, 12).NumberFormat = "0.00%"
                ws.Cells(Summary_row, 13).Value = total_vol
                If per_chg > max_increase Then
                    max_increase = per_chg
                    max_increase_ticker = Ticker_Name
                End If
                If per_chg < max_decrease Then
                    max_decrease = per_chg
                    max_decrease_ticker = Ticker_Name
                End If
                If total_vol > max_volume Then
                    max_volume = total_vol
                    max_volume_ticker = Ticker_Name
                End If
                total_vol = 0
                open_val = 0
                close_val = 0
            Else
                total_vol = total_vol + ws.Cells(i, 7).Value
                If open_val = 0 Then
                    open_val = ws.Cells(i, 3).Value
                End If
            End If
        Next i
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("P2").Value = max_increase_ticker
        ws.Range("Q2").Value = max_increase
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("P3").Value = max_decrease_ticker
        ws.Range("Q3").Value = max_decrease
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P4").Value = max_volume_ticker
        ws.Range("Q4").Value = max_volume
    Next ws
End Sub

