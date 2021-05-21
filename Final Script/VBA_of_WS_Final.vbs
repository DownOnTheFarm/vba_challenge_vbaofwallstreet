Sub stockAnalysis()


For Each ws In Worksheets

    'Set Column Headers Names
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    'Autofit Column Widths
    Worksheets("2014").Range("A:L").Columns.AutoFit
    Worksheets("2015").Range("A:L").Columns.AutoFit
    Worksheets("2016").Range("A:L").Columns.AutoFit
    
    
    'Set Values
    Dim ticker As String
    Dim yearlyOpen As Double
    Dim yearlyClose As Double
    Dim yearlyChange As Double
    Dim totalVolume As Double
    totalVolume = 0
    Dim summaryRow As Long
    summaryRow = 2
    Dim lastRow As Long
    Dim PriorResult As Long
    PriorResult = 2
    Dim GreatestIncrease As Double
    GreatestIncrease = 0
    Dim GreatestDecrease As Double
    GreatestDecrease = 0
    Dim LastRowValue As Long
    Dim GreatestTotalVolume As Double
    GreatestTotalVolume = 0
    
    'Set Last Row
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            
    For i = 2 To lastRow
    
        'For Total Stock Volume
        totalVolume = totalVolume + ws.Cells(i, 7).Value
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        'Set Ticker Name
        ticker = ws.Cells(i, 1).Value
    
        'Print The Ticker Name
        ws.Range("I" & summaryRow).Value = ticker
    
        'Print The Ticker Total Amount
        ws.Range("L" & summaryRow).Value = totalVolume
    
        'Reset Ticker Total
        totalVolume = 0
    
        'Set Yearly Open
        yearlyOpen = ws.Range("C" & PriorResult)
    
        'Set Yearly Close
        yearlyClose = ws.Range("F" & i)
    
        'Set Yearly Change
        yearlyChange = yearlyClose - yearlyOpen
    
        ws.Range("J" & summaryRow).Value = yearlyChange
    
    'Determine Percent Change
    If yearlyOpen = 0 Then
        PercentChange = 0
    Else
        yearlyOpen = ws.Range("C" & PriorResult)
        PercentChange = yearlyChange / yearlyOpen
    End If
    
        'Format Cells to Percentage
        ws.Range("K" & summaryRow).Value = PercentChange
        ws.Range("K" & summaryRow).NumberFormat = "0.00%"
    
    'Highlight Cells- Positive(Green) or Negative(Red)
    If ws.Range("J" & summaryRow).Value >= 0 Then
        ws.Range("J" & summaryRow).Interior.ColorIndex = 4
    Else
        ws.Range("J" & summaryRow).Interior.ColorIndex = 3
    End If

        'Repeat by Adding One To The Summary Table Row
        summaryRow = summaryRow + 1
        PriorResult = i + 1

        End If

    Next i
    
    ' Greatest % Increase, Greatest % Decrease and Greatest Total Volume
            lastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
            ' Start Loop For Final Results
            For i = 2 To lastRow
                If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
                    ws.Range("Q2").Value = ws.Range("K" & i).Value
                    ws.Range("P2").Value = ws.Range("I" & i).Value
                End If

                If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
                    ws.Range("Q3").Value = ws.Range("K" & i).Value
                    ws.Range("P3").Value = ws.Range("I" & i).Value
                End If

                If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
                    ws.Range("Q4").Value = ws.Range("L" & i).Value
                    ws.Range("P4").Value = ws.Range("I" & i).Value
                End If

            Next i
        ' Format Double To Include % Symbol And Two Decimal Places
            ws.Range("Q2").NumberFormat = "0.00%"
            ws.Range("Q3").NumberFormat = "0.00%"
            
        ' Format Table Columns To Auto Fit
        ws.Columns("I:Q").AutoFit


Next ws

End Sub
