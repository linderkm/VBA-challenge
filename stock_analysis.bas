Attribute VB_Name = "Module1"
'This program iterates over all sheets in a workbook. In each sheet, the program iterates over each row of data and records, per ticker, the opening price at the start of the quarter, closing price at the end of the quarter and total volume.
'From the data included in the sheet, the program outputs, per unique ticker, the absolute price change, percentage change and total volume.
'The program outputs the ticker (and value) with the greatest percentage increase, greatest percentage decrease and greatest total volume.

'initializing the function. (1)
Sub stock_analysis()
    
    'initializing ws as an iterable object using the built in Worksheets datatype. This allows the program to iterate over all sheets in the workbook. (2)
    Dim ws As Worksheet
    'Initializing loop. Code inside the loop executes on all sheets in workbook.
    For Each ws In Worksheets
    
        '------------------------------ part 1 of function: parsing raw ticker data -----------------------------------
        'initializing variables targeted in raw data included in sheet.
        Dim ticker As String
        Dim price_qStart As Double
        Dim price_qEnd As Double
        Dim volume As LongLong 'source for datatype (3)
   
        'initializing a variable that identifies which row in the sheet to include data. This prevents the need to hardcode. (2)
        Dim lastrow As Long
        lastrow1 = ws.Cells(Rows.Count, 1).End(xlUp).row
          
        
        'setting initial values in targeted variables
        ticker = ws.Cells(2, 1).Value
        price_qStart = ws.Cells(2, 3).Value
        price_qEnd = 0
        volume = 0
        
        'initializing a variable that will be used in the loop below to control output location in the sheet (i.e. which row will consolidated ticker data be output in).
        Dim report_counter As Long
        repCounter = 2
        
        'formatting column headers for output
        ws.Cells(1, 10).Value = "Ticker"
        ws.Cells(1, 11).Value = "Quarterly Change"
        ws.Cells(1, 12).Value = "Percent Change"
        ws.Cells(1, 13).Value = "Total Stock Volume"
        
        
        'initializing loop that will run over each line of raw ticker data. (4)
        Dim row As Long
        For row = 2 To lastrow1
                
            'conditional statement that check if ticker value matches ticker value the ticker value in the next line. If the next line doesn't match, then the following actions are taken. (5)
            If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then
                'at the row the condition is met, record end of quarter price in the price_qEnd variable, and add the volume value to the total volume (stored in volume variable).
                price_qEnd = ws.Cells(row, 6).Value
                'adding the volume value to the total volume variable.
                volume = volume + ws.Cells(row, 7).Value
                
                'outputting the ticker value, using the output location tracking variable repCounter
                ws.Cells(repCounter, 10).Value = ticker
                
                'outputting the difference between the open price at the start of the quarter (price_qStart) and the closing price at the end of the quarter (price_qEnd)
                ws.Cells(repCounter, 11).Value = (price_qEnd - price_qStart)
                'conditional statement to format the cells, based on the difference between price_qStart and price_qEnd. If > 0, format the cell to green. If < 0, format the cell to red. (6)
                If ws.Cells(repCounter, 11).Value > 0 Then
                    ws.Cells(repCounter, 11).Interior.ColorIndex = 4
                ElseIf ws.Cells(repCounter, 11).Value < 0 Then
                    ws.Cells(repCounter, 11).Interior.ColorIndex = 3
                Else
                End If
                
                'ouputting the difference price_qStart and price_qEnd, as a percentage, as a value without formatting.
                ws.Cells(repCounter, 12).Value = (price_qEnd - price_qStart) / price_qStart
                'formatting the percentage difference as a percentage style type. (7)
                ws.Cells(repCounter, 12).Style = "Percent"
                'conditional statement to format cells. If percentage value > 0, format the cell green. If the percentage value < 0, format the cecll red. (6)
                If ws.Cells(repCounter, 12).Value > 0 Then
                    ws.Cells(repCounter, 12).Interior.ColorIndex = 4
                ElseIf ws.Cells(repCounter, 12).Value < 0 Then
                    ws.Cells(repCounter, 12).Interior.ColorIndex = 3
                Else
                End If
                
                'outputting the total volume.
                ws.Cells(repCounter, 13).Value = volume
                
                'incrementign the output position control variable.
                repCounter = repCounter + 1
                
                'setting the ticker variable to the subsequent unique ticker value (the ticker name in the upcoming row)
                ticker = ws.Cells(row + 1, 1).Value
                'setting the price_qStart variable to the opening price at the start of the quarter for the subsequent unique ticker.
                price_qStart = ws.Cells(row + 1, 3).Value
                'resetting volume to zero
                volume = 0
                
             Else
                'ticker value stays the same.
                ws.Cells(row, 1) = ticker
                'adding the volume value in the current row to the total volume (stored in volume variable).
                volume = volume + ws.Cells(row, 7).Value
    
            End If
            
            
            
        Next row
        
        '------------------------------ part 2 of function: parsing ticker summary--------------------------------------
        'The second primary componenent of the stock_analysis fucntion iterates over the output from part1 of the function.
        'From each ticker summary (output), the code identifies and outputs the ticker with the greatest percentage increase, the greatest precentage decerase, and the greatest total trading volume.
        
        'initializing variables to store data as the coder iterates over the data.
        Dim greatest_increase_ticker As String
        Dim greatest_increase_value As Double
        Dim greatest_decrease_ticker As String
        Dim greatest_decrease_value As Double
        Dim greatest_total_volume_ticker As String
        Dim greatest_total_volume_value As LongLong
        
        'setting a new variable to identify the last row in this specific range. (2)
        Dim lastrow2 As Long
        lastrow2 = ws.Cells(Rows.Count, 10).End(xlUp).row
        
        'setting initial variable values.
        greatest_increase_ticker = " "
        greatest_increase_value = 0
        greatest_decrease_ticker = " "
        greatest_decrease_value = 0
        greatest_total_volume_ticker = " "
        greatest_total_volume_value = 0
        
        
        'conditionl to iterate over rows of data
        For row = 2 To lastrow2
            
            'conditional to check if the cell value is greater than zero, to identify cells with a positive value (candidate for greatest percentage increase).
            If ws.Cells(row, 11).Value > 0 Then
                'if the cell value is greater than the exisitng variable value, then it becomes the new stored value, against which subsequent cell values will be compared. the ticker of the new largest value is also recorded.
                If ws.Cells(row, 12).Value > greatest_increase_value Then
                    greatest_increase_value = ws.Cells(row, 12).Value
                    greatest_increase_ticker = ws.Cells(row, 10).Value
                Else
                End If
            Else
            End If
            
            'conditional to check if the cell value is less than zero, to identify cells with a negative value (candidate for greatest percentage decrease).
            If ws.Cells(row, 11).Value < 0 Then
                'if the cell value is less than the exisitng variable value, then it becomes the new stored value, against which subsequent cell values will be compared. the ticker of the new smallest value is also recorded.
                If ws.Cells(row, 12).Value < greatest_decrease_value Then
                    greatest_decrease_value = ws.Cells(row, 12).Value
                    greatest_decrease_ticker = ws.Cells(row, 10).Value
                Else
                End If
            Else
            End If
                
            
            'conditional to check if the volume is greater than the stored volume. if the cell value is greater than the exisitng variable value, then it becomes the new stored value, against which subsequent cell values will be compared. the ticker of the new largest value is also recorded.
            If ws.Cells(row, 13).Value > greatest_total_volume_value Then
                greatest_total_volume_value = ws.Cells(row, 13).Value
                greatest_total_volume_ticker = ws.Cells(row, 10).Value
            
            Else
            End If
            
        Next row
        
        'formatting column headers for output.
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        'Outputting the ticker with greatest percentage increase (ticker name and value)
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(2, 16).Value = greatest_increase_ticker
        ws.Cells(2, 17).Value = greatest_increase_value
        'formatting the cell to value to show as a percentage. (8)
        ws.Cells(2, 17).Style = "Percent"
        
        'Outputting the ticker with greatest percentage decrease (ticker name and value)
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Value = greatest_decrease_ticker
        ws.Cells(3, 17).Value = greatest_decrease_value
        'formatting the cell to value to show as a percentage. (8)
        ws.Cells(3, 17).Style = "Percent"
        
        'Outputting the ticker with greatest trading volume (ticker name and value)
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(4, 16).Value = greatest_total_volume_ticker
        ws.Cells(4, 17).Value = greatest_total_volume_value
        
        
        
    Next ws
    
    
End Sub



















