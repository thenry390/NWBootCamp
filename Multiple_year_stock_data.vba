Attribute VB_Name = "Module1"
Sub securitiesAnalyzer()

    Dim rowCount As Long
    Dim i As Long
    Dim totalVolume As Variant
    Dim largestDecreaseValue As Double
    Dim largestIncreaseValue As Double
    Dim largestVolume As Variant
    Dim largestDecreaseTicker As String
    Dim largestIncreaseTicker As String
    Dim largestVolumeTicker As String
    Dim ws As Worksheet
    Dim openPrice As Double
    Dim closePrice As Double
    
    For Each ws In ThisWorkbook.Worksheets

        rowCount = 1
        i = 2
        totalVolume = 0
        largestDecreaseValue = 0
        largestIncreaseValue = 0
        largestVolume = 0
        
            
        'populate headers for new columns needed
        ws.Cells(rowCount, 9).Value = "Ticker"
        ws.Cells(rowCount, 10).Value = "Yearly Change"
        ws.Cells(rowCount, 11).Value = "% Change"
        ws.Cells(rowCount, 12).Value = "Total Volume"
        
        Do While ws.Cells(i, 1).Value <> ""
            
            'on change of ticker value, add to rowCount and reset values
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                
                rowCount = rowCount + 1
                ws.Cells(rowCount, 9).Value = ws.Cells(i, 1).Value
                totalVolume = ws.Cells(i, 7).Value
                openPrice = ws.Cells(i, 3).Value
             
            Else 'ticker is not changed, accumulate volume
                
                totalVolume = totalVolume + ws.Cells(i, 7)
            
            End If
            
            If (ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value) Then

                closePrice = ws.Cells(i, 6).Value ' set last closing price
                'determine percent yearly change
                'ws.Cells(rowCount, 10).Value = ws.Cells(i, 3).Value
                ws.Cells(rowCount, 10).Value = closePrice - openPrice
                If openPrice > 0 Then
                    ws.Cells(rowCount, 11).Value = ((ws.Cells(rowCount, 10).Value) / openPrice)
                Else
                    ws.Cells(rowCount, 11).Value = 0
                End If
                
                ws.Cells(rowCount, 11).Value = Format(ws.Cells(rowCount, 11).Value, "0.00%")

                'add conditional shading
                If ws.Cells(rowCount, 10).Value > 0 Then
                    ws.Cells(rowCount, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(rowCount, 10).Interior.ColorIndex = 3
                End If

                ws.Cells(rowCount, 12).Value = totalVolume
            End If
            i = i + 1
       
        Loop
    
        i = 2
        largestDecreaseValue = 0
        largestIncreaseValue = 0
        largestVolume = 0
            
        ' find largest increase, smallest decrease, and largest total volume all for single year
        Do While ws.Cells(i, 9).Value <> ""  'looking for existance of stock ticker, exit if none
        
           If ws.Cells(i, 10).Value > 0 Then
           
                If ws.Cells(i, 10) > largestIncreaseValue Then
                    largestIncreaseValue = ws.Cells(i, 10).Value
                    largestIncreaseTicker = ws.Cells(i, 9).Value
                End If
            Else
                If ws.Cells(i, 10) < largestDecreaseValue Then
                    largestDecreaseValue = ws.Cells(i, 10).Value
                    largestDecreaseTicker = ws.Cells(i, 9).Value
                End If
            End If
        
            If ws.Cells(i, 12) > largestVolume Then
                    largestVolume = ws.Cells(i, 12).Value
                    largestVolumeTicker = ws.Cells(i, 9).Value
            End If
            i = i + 1
       
        Loop
        ' populate values in spreadsheet
        ' start with headers
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        ' next populate three rows of headers and values
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(2, 15).Value = largestIncreaseTicker
        ws.Cells(2, 16).Value = largestIncreaseValue
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(3, 15).Value = largestDecreaseTicker
        ws.Cells(3, 16).Value = largestDecreaseValue
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(4, 15).Value = largestVolumeTicker
        ws.Cells(4, 16).Value = largestVolume
    
    Next ws
   
    
End Sub




