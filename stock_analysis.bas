Attribute VB_Name = "Module1"
Sub stock_analysis()
    
    
    Dim ws As Worksheet
    
    For Each ws In Worksheets
    
        '------------------------------ part 1 of function: parsing raw ticker data
        Dim ticker As String
        Dim price_qStart As Double
        Dim price_qEnd As Double
        Dim volume As LongLong 'source for datatype:https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/data-type-summary
   
        'Set ws = Worksheets("Q1")
        Dim lastrow As Long
        lastrow1 = ws.Cells(Rows.Count, 1).End(xlUp).row
          
        
        Dim row As Long
        'Dim col As Integer
        
        
        
        ticker = ws.Cells(2, 1).Value
        price_qStart = ws.Cells(2, 3).Value
        price_qEnd = 0
        volume = 0
        
        Dim report_counter As Long
        repCounter = 2
        
        ws.Cells(1, 10).Value = "Ticker"
        ws.Cells(1, 11).Value = "Quarterly Change"
        ws.Cells(1, 12).Value = "Percent Change"
        ws.Cells(1, 13).Value = "Total Stock Volume"
        
        
        For row = 2 To lastrow1
                
            
            If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then
                price_qEnd = ws.Cells(row, 6).Value
                volume = volume + ws.Cells(row, 7).Value
                
                ws.Cells(repCounter, 10).Value = ticker
                
                ws.Cells(repCounter, 11).Value = (price_qEnd - price_qStart)
                If ws.Cells(repCounter, 11).Value > 0 Then
                    ws.Cells(repCounter, 11).Interior.ColorIndex = 4
                ElseIf ws.Cells(repCounter, 11).Value < 0 Then
                    ws.Cells(repCounter, 11).Interior.ColorIndex = 3
                Else
                End If
                
                ws.Cells(repCounter, 12).Value = (price_qEnd - price_qStart) / price_qStart
                ws.Cells(repCounter, 12).Style = "Percent" 'source:https://learn.microsoft.com/en-us/office/vba/api/excel.range.style
                If ws.Cells(repCounter, 12).Value > 0 Then
                    ws.Cells(repCounter, 12).Interior.ColorIndex = 4
                ElseIf ws.Cells(repCounter, 12).Value < 0 Then
                    ws.Cells(repCounter, 12).Interior.ColorIndex = 3
                Else
                End If
                
                ws.Cells(repCounter, 13).Value = volume
                
                repCounter = repCounter + 1
                
                ticker = ws.Cells(row + 1, 1).Value
                price_qStart = ws.Cells(row + 1, 3).Value
                volume = 0
                
             Else
                ws.Cells(row, 1) = ticker
            
                volume = volume + ws.Cells(row, 7).Value
    
            End If
            
            
            
        Next row
        
        '------------------------------ part 2 of function: parsing ticker summary
        
        Dim greatest_increase_ticker As String
        Dim greatest_increase_value As Double
        Dim greatest_decrease_ticker As String
        Dim greatest_decrease_value As Double
        Dim greatest_total_volume_ticker As String
        Dim greatest_total_volume_value As LongLong
        
        Dim lastrow2 As Long
        lastrow2 = ws.Cells(Rows.Count, 10).End(xlUp).row
        
        greatest_increase_ticker = " "
        greatest_increase_value = 0
        greatest_decrease_ticker = " "
        greatest_decrease_value = 0
        greatest_total_volume_ticker = " "
        greatest_total_volume_value = 0
        
        
        
        For row = 2 To lastrow2
        
            If ws.Cells(row, 11).Value > 0 Then
            
                If ws.Cells(row, 12).Value > greatest_increase_value Then
                    greatest_increase_value = ws.Cells(row, 12).Value
                    greatest_increase_ticker = ws.Cells(row, 10).Value
                Else
                End If
            Else
            End If
                  
            If ws.Cells(row, 11).Value < 0 Then
                
                If ws.Cells(row, 12).Value < greatest_decrease_value Then
                    greatest_decrease_value = ws.Cells(row, 12).Value
                    greatest_decrease_ticker = ws.Cells(row, 10).Value
                Else
                End If
            Else
            End If
                
            
            
            If ws.Cells(row, 13).Value > greatest_total_volume_value Then
                greatest_total_volume_value = ws.Cells(row, 13).Value
                greatest_total_volume_ticker = ws.Cells(row, 10).Value
            
            Else
            End If
            
        Next row
        
        
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(2, 16).Value = greatest_increase_ticker
        ws.Cells(2, 17).Value = greatest_increase_value
        ws.Cells(2, 17).Style = "Percent"
        
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Value = greatest_decrease_ticker
        ws.Cells(3, 17).Value = greatest_decrease_value
        ws.Cells(3, 17).Style = "Percent"
        
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(4, 16).Value = greatest_total_volume_ticker
        ws.Cells(4, 17).Value = greatest_total_volume_value
        
        
        
    Next ws
    
    
End Sub



















