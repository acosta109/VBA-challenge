Sub stocks():

    Dim ticker As String 'initiailise variable to track ticker
    Dim startPrice As Double 'initialise variable to tracker year start price
    Dim endPrice As Double 'intialise variable to track year end price
    Dim totalVolume As Long 'intiaise variable to track total volume
    Dim summaryTableRow As Integer 'initalise summary row table indicator
    Dim i As Integer 'loop counter
    
    totalVolume = 0 'establish a value now so i can use it right away later
    
    
    For Each ws In Worksheets 'loops through worksheets
    
        'summary rows per worksheet
    
        summaryTableRow = 2 'i want this inside the worksheet for loop at the very top so each time we change worksheets, the variable is reassigned the value 2.
    
        'establish summary row headers
        
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"  'turn green if postive, turn red if negative
        Range("L1").Value = "Total Stock Volume"
        
        
        'Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        startPrice = Range("C2").Value 'i'm initalising this value here so i can change the startPrice only when the ticker value is different. i change this action in the for loop.
        
        For i = 2 To LastRow
        
            'Check if ticker is the same
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
                'print out Ticker
                
                Cells(summaryTableRow, 9).Value = Cells(i, 1).Value
                
                'print out yearly change + read in end price
                
                endPrice = Cells(i, 5).Value
                
                
                Cells(summaryTableRow, 10).Value = endPrice - startPrice
                
                'print out volume
                
                Cell(summaryTableRow, 12) = totalVolume
                totalVolume = 0 ' reset totalVolume
                
            Else 'adds volume because the stock is the same
                totalVolume = CLng(totalVolume) + Cells(i, 7).Value
                
            End If
                 
        Next i
        
    Next ws 'goes to next worksheet

End Sub
