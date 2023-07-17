VBA-challenge
Sub multi_year_stock()
    For Each ws In Worksheets
        ws.Activate

        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        
        Range("O2").Value = "Greatest Percent Increase"
        Range("O3").Value = "Greatest Percent Decrease"
        Range("O4").Value = "Greatest Total Stock Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        
        Columns("A:Q").AutoFit
        
            ' Set an initial variable for holding the brand name
            Dim ticker As String

            ' Set an initial variable for holding the total per credit card brand
            Dim totalVol As Double
            totalVol = 0

            ' Keep track of the location for each credit card brand in the summary table
            Dim summary_pointer As Integer
            summary_pointer = 2
            openPricePointer = 2
        
            ' Find the last row of data in the worksheet
            lastRow = Cells(Rows.Count, "A").End(xlUp).Row

            ' Loop through all credit card purchases
             For i = 2 To lastRow

            ' Check if we are still within the same credit card brand, if it is not...
                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                    totalVol = totalVol + Cells(i, "G").Value
                    openPrice = Cells(openPricePointer, "C").Value
                    closePrice = Cells(i, "F").Value
                    
                    yearlyChange = closePrice - openPrice
                    percentageChange = yearlyChange / openPrice * 100
                    
                    Cells(summary_pointer, "I").Value = Cells(i, "A").Value
                    Cells(summary_pointer, "J").Value = yearlyChange
                    Cells(summary_pointer, "K").Value = "%" & percentageChange
                    Cells(summary_pointer, "L").Value = totalVol
                    
                    If yearlyChange > 0 Then
                        Range("J" & summary_pointer).Interior.Color = vbGreen
                    ElseIf yearlyChange < 0 Then
                        Range("J" & summary_pointer).Interior.Color = vbRed
                    Else
                        Range("J" & summary_pointer).Interior.Color = vbWhite
                    End If
                    
                    totalVol = 0
                    summary_pointer = summary_pointer + 1
                    openPricePointer = i + 1
                    
                Else

                    ' Add to the Total
                    totalVol = totalVol + Cells(i, "G").Value

                End If

            Next i
            
            Dim greatest_percent_increase As Double
            Dim greatest_percent_decrease As Double
            Dim greatest_total As Long
            Dim ticker_name As String
                
            greatest_percent_increase = 0
            greatest_percent_decrease = 0
            greatest_total_volume = 0
            
            summary_count = Cells(Rows.Count, "K").End(xlUp).Row
            
            'List the Greatest Percentate Increase
            
            For Row = 2 To summary_count
            'REM- Compare to find Greatest_percent_Increase
                If Cells(Row, "K").Value > greatest_percent_increase Then
                    'MsgBox (Cells(row, "K").Value)
                    greatest_percent_increase = Cells(Row, "K").Value
                    ticker_name = Cells(Row, "I").Value
                End If
            Next Row

            Range("P2") = ticker_name
            Range("Q2") = greatest_percent_increase
            
            'List the Greatest Percentate Decrease
            
            For Row = 2 To summary_count
            'REM- Compare to find Greatest_percent_decrease
                If Cells(Row, "K").Value < greatest_percent_decrease Then
                    'MsgBox (Cells(row, "K").Value)
                    greatest_percent_decrease = Cells(Row, "K").Value
                    ticker_name = Cells(Row, "I").Value
                End If
            Next Row

            Range("P3") = ticker_name
            Range("Q3") = greatest_percent_decrease
            
            'List the Total Stock Volume
            
            For Row = 2 To summary_count
            'REM- Compare to find Greatest_percent_Increase
                If Cells(Row, "K").Value > greatest_total_volume Then
                    'MsgBox (Cells(row, "K").Value)
                    greatest_total_volume = Cells(Row, "K").Value
                    ticker_name = Cells(Row, "I").Value
                End If
            Next Row

            Range("P4") = ticker_name
            Range("Q4") = greatest_total_volume
            
            Columns("A:Q").AutoFit
            
    Next ws
    
End Sub