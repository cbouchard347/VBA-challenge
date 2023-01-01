Attribute VB_Name = "Module1"
Sub HW2()

  For Each ws In Worksheets
    
        Dim WorksheetName As String
        Dim i As Long
        Dim j As Long
        Dim Ticker_Count As Long
        Dim LastRowA As Long
        Dim LastRowI As Long
        Dim Precent_Change As Double
        Dim Greatest_Percent_Increase As Double
        Dim Greatest_Percent_Decrease As Double
        Dim Greatest_Volume As Double
        
        WorksheetName = ws.Name
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(1, 15).Value = "Greatest % Increase"
        ws.Cells(2, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 15).Value = "Greatest Total Volume"
                
        Ticker_Count = 2
        
        j = 2
        
        LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
            For i = 2 To LastRowA
            
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                ws.Cells(Ticker_Count, 9).Value = ws.Cells(i, 1).Value
                
                ws.Cells(Ticker_Count, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                    If ws.Cells(Ticker_Count, 10).Value < 0 Then
                    ws.Cells(Ticker_Count, 10).Interior.ColorIndex = 3
                
                    Else
                    ws.Cells(Ticker_Count, 10).Interior.ColorIndex = 4
                
                    End If
                    
                    If ws.Cells(j, 3).Value <> 0 Then
                    Percent_Change = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    
                    ws.Cells(Ticker_Count, 11).Value = Format(Percent_Change, "Percent")
                    
                    Else
                    
                    ws.Cells(Ticker_Count, 11).Value = Format(0, "Percent")
                    
                    End If

                ws.Cells(Ticker_Count, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))

                Ticker_Count = Ticker_Count + 1

                j = i + 1
                
                End If
            
            Next i
            
        LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        Greatest_Volume = ws.Cells(2, 12).Value
        Greatest_Percent_Increase = ws.Cells(2, 11).Value
        Greatest_Percent_Decrease = ws.Cells(2, 11).Value
        
            For i = 2 To LastRowI
            
                If ws.Cells(i, 12).Value > Greatest_Volume Then
                Greatest_Volume = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                Greatest_Volume = Greatest_Volume
                
                End If
                
                If ws.Cells(i, 11).Value > Greatest_Percent_Increase Then
                Greatest_Percent_Increase = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                Greatest_Percent_Increase = Greatest_Percent_Increase
                
                End If
                
                If ws.Cells(i, 11).Value < Greatest_Percent_Decrease Then
                Greatest_Percent_Decrease = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                Greatest_Percent_Decrease = Greatest_Percent_Decrease
                
                End If

            ws.Cells(2, 17).Value = Format(Greatest_Percent_Increase, "Percent")
            ws.Cells(3, 17).Value = Format(Greatest_Percent_Decrease, "Percent")
            ws.Cells(4, 17).Value = Format(Greatest_Volume, "Scientific")
            
            Next i
            
    Next ws
        
End Sub

