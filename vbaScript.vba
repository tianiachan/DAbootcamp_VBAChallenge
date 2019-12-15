Sub alphaStock()
    For Each ws In Worksheets
        'store values
        Dim tickerValue As String
        Dim yearlyChange As Long
        Dim percentChange As Double
        Dim totalStock As Double
        Dim openVal As Long
        Dim closeVal As Long
        
        Range("I1", "L1").WrapText = True
        Range("I1", "L1").Font.Bold = True
        'set ticker header
        Cells(1, 9).Value = "Ticker"
        'set percent change header
        Cells(1, 10).Value = "Percent Change"
        'set yearly change header
        Cells(1, 11).Value = "Yearly Change"
        'set Total stock header
        Cells(1, 12).Value = "Total Stock Volume"
        
        'Summary row value
        Dim sumRow As Integer
        sumRow = 2
        'get length of row
        rowLen = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To rowLen + 1:
        ' if not same ticker, print everything in summary
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                'print ticker value
                Cells(sumRow, 9).Value = Cells(i, 1).Value
                'print totalValue
                Cells(sumRow, 12).Value = totalStock
                'print yearlyChange
                yearlyChange = closeVal - openVal
                Cells(sumRow, 11).Value = yearlyChange
                'print percentChange
                percentChange = ((closeVal - openVal) / openVal) * 100
                'red if negative change, green of positive change
                If percentChange > 0 Then
                    Cells(sumRow, 10).Interior.Color = vbGreen
                Else
                    Cells(sumRow, 10).Interior.Color = vbRed
                End If
                Cells(sumRow, 10).Value = percentChange
                
                'move onto next row and reset all values
                sumRow = sumRow + 1
                openVal = 0
                closeVal = 0
                totalStock = 0
                percentChange = 0
                
            'if not do calculations and save values in variables
            Else
                'add to total
                totalStock = totalStock + Cells(i, 7).Value
                'add to openVal
                openVal = openVal + Cells(i, 3).Value
                'add to closeVal
                closeVal = closeVal + Cells(i, 6).Value
            End If
        Next i
    Next ws
End Sub

