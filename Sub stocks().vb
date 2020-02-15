
Sub stocks()

        For Each ws In Worksheets

            'Defining variables
             
            Dim i As Long
            Dim TickerSymbol As String
            Dim YearlyChange As Double
            Dim StockVolume As LongLong
            Dim NewDataRow As Integer
            Dim YearOpen As Double
            Dim YearClose As Double
            Dim GreatestIncrease As Double
            Dim GreatestDecrease As Double
            Dim GreatestVolume As LongLong
            
            'NewDataRow adds new rows on the summary table
            NewDataRow = 2
            StockVolume = 0
            GreatestVolume = 0
            GreatestIncrease = 0
            GreatestDecrease = 1

            LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

                'Headers for the Summary Tables 
                ws.Range("I1").Value = "Ticker"
                ws.Range("J1").Value = "Yearly Change"
                ws.Range("K1").Value = "Percent Change"
                ws.Range("L1").Value = "Total Stock Volume"

                ws.Range("O2").Value = "Greatest % Increase"
                ws.Range("O3").Value = "Greatest % Decrease"
                ws.Range("O4").Value = "Greatest Total Volume"
                ws.Range("P1").Value = "Ticker"
                ws.Range("Q1").Value = "Value"
                
                    'For loop to count through all of the rows in the sheet
                    For i = 2 To LastRow
                            
                            'If the following cell has a different ticker than the current cell, store the current ticker,
                            ' add stock volume, store the year close, then calculate the total yearly change
                            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                                TickerSymbol = ws.Cells(i, 1).Value
                                StockVolume = StockVolume + ws.Cells(i, 7).Value
                                YearClose = ws.Cells(i, 6).Value
                                YearlyChange = YearClose - YearOpen

                                'Print out the stored values on to the new summary table
                                ws.Range("I" & NewDataRow).Value = TickerSymbol
                                ws.Range("J" & NewDataRow).Value = YearlyChange

                                If YearOpen > 0 Then
                                    ws.Range("K" & NewDataRow).Value = FormatPercent((YearlyChange / YearOpen))
                                Else
                                    ws.Range("K" & NewDataRow).Value = FormatPercent(0)
                                End If

                                ws.Range("L" & NewDataRow).Value = StockVolume

                                    'Conditional formatting for the % change column
                                    If ws.Range("J" & NewDataRow).Value > 0 Then
                                        ws.Range("J" & NewDataRow).Interior.ColorIndex = 4
                                    ElseIf ws.Range("J" & NewDataRow) < 0 Then
                                        ws.Range("J" & NewDataRow).Interior.ColorIndex = 3
                                    Else
                                        ws.Range("J" & NewDataRow).Interior.ColorIndex = 6
                                    End If
                                    
                                    If ws.Range("K" & NewDataRow).Value > GreatestIncrease Then
                                        GreatestIncrease = ws.Range("K" & NewDataRow).Value
                                        ws.Range("P2").Value = TickerSymbol
                                        ws.Range("Q2").Value = FormatPercent(GreatestIncrease)
                                    End If
                                    If ws.Range("K" & NewDataRow).Value < GreatestDecrease Then
                                        GreatestDecrease = ws.Range("K" & NewDataRow).Value
                                        ws.Range("P3").Value = TickerSymbol
                                        ws.Range("Q3").Value = FormatPercent(GreatestDecrease)
                                    End If
                                    If ws.Range("L" & NewDataRow).Value > GreatestVolume Then
                                        GreatestVolume = ws.Range("L" & NewDataRow).Value
                                        ws.Range("P4").Value = TickerSymbol
                                        ws.Range("Q4").Value = StockVolume
                                    
                                    End If

                                NewDataRow = NewDataRow + 1
                                StockVolume = 0
                                YearlyChange = 0
                                
                            ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                                YearOpen = ws.Cells(i, 3).Value

                            'If the current ticker is the same as both of the adjacent tickers, add stock volume   
                            Else
                                StockVolume = StockVolume + ws.Cells(i, 7).Value
                                
                            End If


                    Next i
                
         Next ws   

         
End Sub


