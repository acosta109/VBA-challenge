Sub Stocks()

    'variable to track ticker
    Dim ticker As String
    'variable to tracker year start price
    Dim startPrice As Double
    'variable to track year end price
    Dim endPrice As Double
    'variable to track total volume
    Dim totalVolume As Double
    'summary row table indicator
    Dim summaryTableRow As Integer
    'loop counter
    Dim i As Double
    'last row counter
    Dim LastRow As Double
    'worksheet to travel
    Dim ws As Worksheet
    'greatet percent decrease
    Dim greatestPerDec As Double
    'greatest perfect increase
    Dim greatestPerInc As Double
    'largest vol
    Dim largVol As Double


    
    For Each ws In ThisWorkbook.Worksheets
        largVol = 0
        'i intialise a value so i can use this in an equation.
        totalVolume = 0

        'summary row starts in row 2 in every workseet
        summaryTableRow = 2
        
        'establish summary row headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        'turn green if postive, turn red if negative
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volue"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"
    
        'Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
   
            For i = 2 To LastRow
        
            'Check if ticker is the same
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                'need to add the volume for the last day.
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                
                'display and reset total volu
                ws.Cells(summaryTableRow, 12) = totalVolume

                'check for starting value, print out the volume and ticker as it's the largest at this point in time

                If ws.Range("Q4") = "" Then
                    largVol = totalVolume
                    ws.Range("Q4").Value = largVol
                    ws.Range("P4").Value = ws.Cells(i, 1).Value
                End If 

                'if the new total volume is greater than the preview largest volume, print out the new information
                If totalVolume > ws.Range("Q4").Value Then
                    largVol = totalVolume
                    ws.Range("Q4").Value = largVol
                    ws.Range("P4").Value = ws.Cells(i, 1).Value
                End If 


                totalVolume = 0
                
                'print out Ticker
                ws.Cells(summaryTableRow, 9).Value = ws.Cells(i, 1).Value
                
                'print out yearly change + read in end price
                endPrice = Cells(i, 5).Value
                ws.Cells(summaryTableRow, 10).Value = endPrice - startPrice
                
                'change colour
                If (endPrice - startPrice) < 0 Then
                    ws.Cells(summaryTableRow, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(summaryTableRow, 10).Interior.ColorIndex = 4
                
                End If
                
                'print out percent change
                ws.Cells(summaryTableRow, 11).Value = (endPrice - startPrice) / startPrice
                ws.Cells(summaryTableRow, 11).NumberFormat = "0.00%"

                'test is this is a new work sheet and estbalish starting values
                If ws.Range("Q2").Value = "" Then
                    greatestPerDec = (endPrice - startPrice) / startPrice
                    greatestPerInc = (endPrice - startPrice) / startPrice
                    ws.Range("Q2").Value = greatestPerInc
                    ws.Range("Q3").Value = greatestPerDec
                    ws.Range("P2").Value = ws.Cells(i, 1).Value
                    ws.Range("P3").Value = ws.Cells(i, 1).value
                End If

                If (endPrice - startPrice) / startPrice > greatestPerInc Then
                    greatestPerInc = (endPrice - startPrice) / startPrice
                    ws.Range("Q2").Value = greatestPerInc
                    ws.Range("P2").Value = ws.Cells(i, 1).Value
                End If

                If (endPrice - startPrice) / startPrice < greatestPerDec Then
                    greatestPerDec = (endPrice - startPrice) / startPrice
                    ws.Range("Q3").Value = greatestPerDec
                    ws.Range("P3").Value = ws.Cells(i, 1).value
                End If
                summaryTableRow = summaryTableRow + 1
            'adds volume because the stock is the same
            Else 
                'track start price
                If totalVolume = 0 Then
                    startPrice = ws.Cells(i, 3).Value
                End If
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                            

            End If
            
        
        Next i
        
    
    Next ws
    

End Sub



