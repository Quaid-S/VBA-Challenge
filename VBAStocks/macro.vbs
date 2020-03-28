Sub analysis():

    Dim totalChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim i As Long
    Dim j As Integer
    Dim firstValue As Long
    Dim lastRow As Long
    Dim ws As Worksheet
    
    For Each ws In Worksheets
        ' Set values for each worksheet
        j = 0
        totalVolume = 0
        totalChange = 0
        firstValue = 2
        

        ' create titles on row 1
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Total Change"
        ws.Range("K1").Value = "Percentage Change"
        ws.Range("L1").Value = "Total Stock Volume"

        ' get the row number of the last row with data
        lastRow = Cells(Rows.Count, "A").End(xlUp).Row

        For i = 2 To lastRow

            ' Handler for ticker changes
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                totalVolume = totalVolume + ws.Cells(i, 7).Value

                ' When Volume = 0
                If totalVolume = 0 Then
                    ws.Range("I" & 2 + j).Value = Cells(i, 1).Value
                    ws.Range("J" & 2 + j).Value = 0
                    ws.Range("K" & 2 + j).Value = "%" & 0
                    ws.Range("L" & 2 + j).Value = 0

                Else
                'starting values of price
                    If ws.Cells(firstValue, 3) = 0 Then
                        For x = firstValue To i
                            If ws.Cells(x, 3).Value <> 0 Then
                                firstValue = x
                                Exit For
                            End If
                        Next x
                    End If

                    ' Total Change in stock value
                    totalChange = Round((ws.Cells(i, 6) - ws.Cells(firstValue, 3)), 2)
                    percentChange = Round((totalChange / ws.Cells(firstValue, 3) * 100), 4)

                    firstValue = i + 1

                    ' print the results to a separate worksheet
                    ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                    ws.Range("J" & 2 + j).Value = totalChange
                    ws.Range("K" & 2 + j).Value = "%" & percentChange
                    ws.Range("L" & 2 + j).Value = totalVolume

                    ' Conditional for colors based on +/- percent change
                    If totalChange > 0 Then
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 4
                    ElseIf totalChange < 0 Then
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 3
                    Else
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 0
                    End If
                End If
                'Reset Variables
                totalVolume = 0
                totalChange = 0
                j = j + 1


            
            Else
                totalVolume = totalVolume + ws.Cells(i, 7).Value

            End If

        Next i

    Next ws

End Sub