# vba_challenge
Sub stock_market()


Dim Ticker As String
Dim Ticker_Row As Integer
Dim total_stock_share As Double
Dim percent_change As Double
Dim Greatest_decrease As Double
Dim yearly_change As Double
Dim Greatest_Increase As Double
Dim Greatest_total_volume As Double
Dim open_value As Double
Dim ws As Worksheet

  For Each ws In Worksheets
      total_stock_share = 0
      Ticker_Row = 2
      open_value = ws.Cells(2, 3).Value
     

      '############"Ticker", "Yearly Change", "Percent change" and "Total Stock Volume"##############

       
       ws.Cells(1, 9).Value = "Ticker"
      ws.Cells(1, 11).Value = "Percent Change"
      ws.Cells(1, 10).Value = "Yearly Change"
      ws.Cells(1, 12).Value = "Total Stock Volume"


            '###############"Ticker" and "Total Stock Volume"##############
            
            LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
                For i = 2 To LastRow
      
       
                    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
         
                         Ticker = ws.Cells(i, 1).Value
             
                         total_stock_share = total_stock_share + ws.Cells(i, 7).Value
                         
                          ws.Range("I" & Ticker_Row).Value = Ticker
                          
                          ws.Range("L" & Ticker_Row).Value = total_stock_share
                          
           '###########computing yearly change and percent change,#############
                        
             
                         yearly_change = ws.Cells(i, 6) - open_value
                         ws.Range("J" & Ticker_Row).Value = yearly_change
                         
                            If open_value = 0 Then
                               percent_change = 0
                             Else
                           
                            percent_change = yearly_change / open_value
                            End If
                         
                         ws.Range("K" & Ticker_Row).Value = percent_change
             
            
             
                            Ticker_Row = Ticker_Row + 1
             
                             total_stock_share = 0
                             
                             open_value = ws.Cells(1 + i, 3)
                            
                    Else

                            total_stock_share = total_stock_share + ws.Cells(i, 7).Value
          
                    End If
                      
        
                Next i
'##############Conditional formatting ###################
                       LR_Yearly_change = ws.Cells(Rows.Count, 10).End(xlUp).Row
                      
                      For R = 2 To LR_Yearly_change
                          If ws.Cells(R, 10).Value < 0 Then
                             ws.Cells(R, 10).Interior.ColorIndex = 3
                             
                          ElseIf ws.Cells(R, 10).Value > 0 Then
                             ws.Cells(R, 10).Interior.ColorIndex = 10
                          
                          End If
                          
                      Next R
                      
'###############correct the percent change to %####################
                      
                      For K = 2 To LR_Yearly_change
                        ws.Range("k2:K" & LR_Yearly_change).NumberFormat = "0.00%"
                      Next K
                      

                      
                        ws.Cells(2, 15).Value = "Greatest % Increase"
                        ws.Cells(3, 15).Value = "Greatest % Decrease"
                        ws.Cells(4, 15).Value = "Greatest Total Volume"
                        ws.Cells(1, 16).Value = "Ticker"
                        ws.Cells(1, 17).Value = "Value"
                        
                        
                        
                        LR_Percent_change = ws.Cells(Rows.Count, 11).End(xlUp).Row
                        For s = 2 To LR_Percent_change
                        
                          '##########determine the max and min percentage change#############
                          
                          If ws.Cells(s, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & LR_Percent_change)) Then
                          
                             ws.Cells(2, 17).Value = ws.Cells(s, 11).Value
                             ws.Cells(2, 16).Value = ws.Cells(s, 9).Value
                             ws.Range("Q3").NumberFormat = "0.00%"
                          
                          ElseIf ws.Cells(s, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & LR_Percent_change)) Then
                        
                             ws.Cells(3, 17).Value = ws.Cells(s, 11).Value
                             ws.Cells(3, 16).Value = ws.Cells(s, 9).Value
                             ws.Range("Q2").NumberFormat = "0.00%"
                          
                           '###########determine the maximum total stock volume#############
                           
                          ElseIf ws.Cells(s, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & LR_Percent_change)) Then
                          
                             ws.Cells(4, 17).Value = ws.Cells(s, 12).Value
                             ws.Cells(4, 16).Value = ws.Cells(s, 9).Value
                             
                          End If
                          
                        Next s

        Next ws


End Sub
