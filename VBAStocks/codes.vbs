Sub homework():

Dim lastrow, lastrow1, total, i, a, b, total_count, x As LongLong


Dim GP_Increase, GP_Decrease, closing_value, opening_value As Double

    For Each ws In Worksheets
        
        ws.Cells(1, 10).Value = "Ticker"
        ws.Cells(1, 11).Value = "Yearly Change"
        ws.Cells(1, 12).Value = "Percentage Change"
        ws.Cells(1, 13).Value = "Total Stock Volume"
        ws.Cells(2, 15).Value = "Greatest Percent Increase"
        ws.Cells(3, 15).Value = "Greatest Percent Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        
        i = 2
        a = 2
        b = 2
        total = 0
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'MsgBox (lastrow)
        lastrow = lastrow + 1
        
            Do While i < lastrow
                
                total = total + ws.Cells(i, 7).Value
                
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    ws.Cells(a, 10).Value = ws.Cells(i, 1).Value
                    opening_value = ws.Cells(b, 3).Value
                    
                    closing_value = ws.Cells(i, 6).Value
                    ws.Cells(a, 11).Value = (closing_value) - (opening_value)
                    If opening_value = 0 Then
                        opening_value = 1
                        End If
                        
                    ws.Cells(a, 12).Value = (ws.Cells(a, 11).Value) / (opening_value)
                    
                    ws.Cells(a, 13).Value = total
                    total = 0
                    b = i + 1
                    'MsgBox ("first opening value of second series is" + Str(b))
                    a = a + 1
                    
                    
                End If
                    
                i = i + 1
            
            Loop
            
          lastrow1 = ws.Cells(Rows.Count, 12).End(xlUp).Row
          GP_Increase = ws.Cells(2, 12).Value
         GP_Decrease = ws.Cells(2, 12).Value
         lastrow1 = lastrow1 + 1
         x = 2
         total_count = ws.Cells(2, 13).Value
         
          
          Do While x < lastrow1
            
            
            
                If GP_Increase < ws.Cells((x + 1), 12).Value Then
                    GP_Increase = ws.Cells((x + 1), 12).Value
                End If
                  
                If GP_Decrease > ws.Cells((x + 1), 12).Value Then
                    GP_Decrease = ws.Cells((x + 1), 12).Value
                End If
                
                If total_count < ws.Cells((x + 1), 13).Value Then
                     total_count = ws.Cells((x + 1), 13).Value
                End If
                                    
                If ws.Cells(x, 12).Value > 0 Then
                    ws.Cells(x, 12).Interior.ColorIndex = 4
                End If
                
                If ws.Cells(x, 12).Value < 0 Then
                    ws.Cells(x, 12).Interior.ColorIndex = 3
                End If
                
                ws.Cells(x, 12).Style = "Percent"
                x = x + 1
        Loop
        
          
            ws.Cells(2, 16).Value = GP_Increase
            ws.Cells(2, 16).Style = "Percent"
            ws.Cells(3, 16).Value = GP_Decrease
            ws.Cells(3, 16).Style = "Percent"
            ws.Cells(4, 16).Value = total_count
              
    Next ws
    
    End Sub


