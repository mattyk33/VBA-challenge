Attribute VB_Name = "Module1"
Sub Stock_Market()

    For Each ws In Worksheets
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        Dim Ticker As String
        Dim YearlyOpen As Double
        Dim YearlyClose As Double
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim TotalVol As Double
        Dim Summary_Table_Row As Long
        Dim LastRow As Long
      
        TotalVol = 0
        Summary_Table_Row = 2
        
        YearlyOpen = ws.Cells(2, 3).Value
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        
        For i = 2 To LastRow
        
            ' Condtional for new ticker
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
                
                ' Ticker and Volume
                Ticker = ws.Cells(i, 1).Value
                TotalVol = TotalVol + ws.Cells(i, 7).Value

                ws.Range("I" & Summary_Table_Row).Value = Ticker
                ws.Range("L" & Summary_Table_Row).Value = TotalVol
                
                ' Yearly Change
                YearlyClose = ws.Cells(i, 6)
                YearlyChange = YearlyClose - YearlyOpen
                ws.Range("J" & Summary_Table_Row).Value = YearlyChange
                
                ' Conditional to find Percent Change
                If (YearlyOpen = 0 And YearlyClose = 0) Then
                    PercentChange = 0
                    
                ElseIf (YearlyOpen = 0 And YearlyClose <> 0) Then
                    PercentChange = Null
                
                Else
                    PercentChange = (YearlyChange / YearlyOpen)
                    ws.Range("K" & Summary_Table_Row).Value = PercentChange
                    ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                    
                End If
            
                Summary_Table_Row = Summary_Table_Row + 1
            
                TotalVol = 0
                
                YearlyOpen = ws.Cells(i + 1, 3).Value
                
            Else
                TotalVol = TotalVol + ws.Cells(i, 7).Value
            
            End If
            
            'Conditional highlight positive/negative
            If ws.Range("J" & Summary_Table_Row).Value > 0 Then
                
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                
            ElseIf ws.Range("J" & Summary_Table_Row).Value < 0 Then
            
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                
            End If
            
        Next i
        
        ' Find "Greatest % increase", "Greatest % decrease" and "Greatest total volume"
        Dim CLastRow As Long
        Dim MaxInc As Double
        Dim MinDec As Double
        Dim MaxVol As Double
        
        MaxInc = 0
        MinDec = 0
        MaxVol = 0

        CLastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row

        For j = 2 To CLastRow
            
            If ws.Range("K" & j).Value > MaxInc Then
                MaxInc = ws.Range("K" & j).Value
                ws.Range("Q2").Value = MaxInc
                ws.Range("Q2").NumberFormat = "0.00%"
                ws.Range("P2").Value = ws.Range("I" & j).Value
                        
            ElseIf ws.Range("K" & j).Value < MinDec Then
                MinDec = ws.Range("K" & j).Value
                ws.Range("Q3").Value = MinDec
                ws.Range("Q3").NumberFormat = "0.00%"
                ws.Range("P3").Value = ws.Range("I" & j).Value

            End If
            
            If ws.Range("L" & j).Value > MaxVol Then
                MaxVol = ws.Range("L" & j).Value
                ws.Range("Q4").Value = MaxVol
                ws.Range("P4").Value = ws.Range("I" & j).Value
            
            End If
            
        Next j
        
    Next ws
    
End Sub

