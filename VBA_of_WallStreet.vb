Sub Stock_Summary()
  

    For Each ws In Worksheets
    
        ' Set an initial variable for holding the Ticker
        Dim Ticker As String

        ' Set an initial variable for holding the total per Ticker
        Dim Ticker_Total As Double
        Ticker_Total = 0
  
        'count num of repetitions for each Ticker
        Dim Counter As Long
        Counter = 0
  
        Dim Yearly_Diff As Double
        Yearly_Diff = 0
  
        Dim Percent_Change As Double
        Percent_Change = 0

        Dim Summary_Table_Row As Long
        Summary_Table_Row = 2

        'Summary table headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
        ' Loop through all rows on each sheet
        For i = 2 To LastRow

            ' Check if we are still within the same credit card brand, if it is not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                ' Set the Ticker
                Ticker = ws.Cells(i, 1).Value

                ' Calculate
                
                
            
                If ws.Cells(i - Counter, 3) <> 0 Then
                    
                    Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
                    
                    Yearly_Diff = ws.Cells(i, 6).Value - ws.Cells(i - Counter, 3).Value
                
                    Percent_Change = Yearly_Diff / ws.Cells(i - Counter, 3).Value
                Else
                    
                    Ticker_Total = 0
                    Yearly_Diff = 0
                    Percent_Change = 0

                End If

                ' populate summary table
                ws.Range("I" & Summary_Table_Row).Value = Ticker

                ws.Range("L" & Summary_Table_Row).Value = Ticker_Total
            
                ' print count of ticker
                'Range("L" & Summary_Table_Row).Value = Counter
               
                ws.Range("J" & Summary_Table_Row).Value = Yearly_Diff
                
                ws.Range("K" & Summary_Table_Row).Value = Percent_Change
                
                        
                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
            
                ' Reset the Brand Total
                Ticker_Total = 0
            
                ' Reset the counter
                Counter = 0

                ' If the cell immediately following a row is the same brand...
            Else

                ' Add to the Ticker Total
                Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
        
                Counter = Counter + 1

            End If
        
        Next i
        
        'Reinitialize
        'Summary_Table_Row = 2
        
        'Format summary stable
        'Find last row of summary table
        LastRowSummary = ws.Cells(Rows.Count, "I").End(xlUp).Row

        For i = 2 To LastRowSummary
           
            If ws.Cells(i, "J") < 0 Then
                ws.Cells(i, "J").Interior.ColorIndex = 3
            Else
                ws.Cells(i, "J").Interior.ColorIndex = 4
            End If

            ws.Cells(i, "K").NumberFormat = "0.00%"

        Next i
        
        'Find Max of each sheet
        Dim MaxVolume As Double
        Dim MaxVolume_Ticker As String
        MaxVolume = 0
        
        Dim MaxPercentUp As Double
        Dim MaxPercentUp_Ticker As String
        MaxPercentUp = 0
        
        Dim MaxPercentDown As Double
        Dim MaxPercentDown_Ticker As String
        MaxPercentDown = 0
        
        For i = 2 To LastRowSummary
        
            If ws.Cells(i, "L").Value > MaxVolume Then
            
                MaxVolume = ws.Cells(i, "L").Value
                MaxVolume_Ticker = ws.Cells(i, "I").Value
                
            End If
            
            If ws.Cells(i, "K").Value > MaxPercentUp Then
            
                MaxPercentUp = ws.Cells(i, "K").Value
                MaxPercentUp_Ticker = ws.Cells(i, "I").Value
                
            End If
            
            If ws.Cells(i, "K").Value < MaxPercentDown Then
                
                MaxPercentDown = ws.Cells(i, "K").Value
                MaxPercentDown_Ticker = ws.Cells(i, "I").Value
            
            End If
                    
        Next i
        
        'Check
        'Output summary of Max for each sheet
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        ws.Range("P2").Value = MaxPercentUp_Ticker
        ws.Range("P3").Value = MaxPercentDown_Ticker
        ws.Range("P4").Value = MaxVolume_Ticker
               
        ws.Range("Q2").Value = MaxPercentUp
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").Value = MaxPercentDown
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("Q4").Value = MaxVolume
        
    Next ws
    
End Sub