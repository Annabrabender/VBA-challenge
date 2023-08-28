Attribute VB_Name = "Module1"
Sub Mulitipleyearstockdata():

For Each ws In Worksheets


    'Set the variables for everything I will need
    
    Dim WorksheetName As String
    Dim i As Long
    Dim j As Long
    
    
    
    Dim LastrowI As Long
    
    
    Dim PercentChange As Double
    Dim Greatestincrease As Double
    Dim Greatestdecrease As Double
    Dim Greatestvolume As Double
    
    
    'Need to establish what the worksheet name is
    WorksheetName = ws.Name
    
    
    'Column headers for Ticker, Yearly Change, Percent Chnage and Total Stock Volume
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    
    
    
    
    'Establish where I want ticker counter to begin
    Dim TickerCount As Long
    TickerCount = 2
    
   
    
    
    'Use last row forumula used in class to establish where the ticker count can stop
    
    Dim LastrowA As Long
    LastrowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        'Need to make sure it is looping all rows
            For i = 2 To LastrowA
                
                'Use formula to check if the Ticker sequence changed
                    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    
                    'Put the ticker statement into the Ticker coloumn
                        ws.Cells(TickerCount, 9).Value = ws.Cells(i, 1).Value
                        
                        
                    'Next for this row we need to calculate the yearly change
                        ws.Cells(TickerCount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(i, 3).Value
                        
                    
                    
                        'Conditional formatting for yearlychange
                        
                            'Set green for positive
                                If ws.Cells(TickerCount, 10).Value > 0 Then
                                ws.Cells(TickerCount, 10).Interior.ColorIndex = 4
                                
                                
                            'Set the rest red for negative
                            
                                Else
                                
                                ws.Cells(TickerCount, 10).Interior.ColorIndex = 3
                                
                                End If
                                
                                
                                
                        ' Now need to code for Percent Change in the next coloumn
                        
                            If ws.Cells(i, 3).Value <> 0 Then
                            PercentChange = ((ws.Cells(i, 6).Value - ws.Cells(i, 3).Value) / ws.Cells(i, 3).Value)
                            
                            'Need to tell it to show us in % format
                                ws.Cells(TickerCount, 11).Value = Format(PercentChange, "Percent")
                                
                            
                            Else
                            
                            ws.Cells(TickerCount, 11).Value = Format(0, "Percent")
                            
                            End If
                            
                            
                        'Conditional formatting for percentchange
                        
                            'Set green for positive
                                If ws.Cells(TickerCount, 11).Value > 0 Then
                                ws.Cells(TickerCount, 11).Interior.ColorIndex = 4
                                
                                
                            'Set the rest red for negative
                            
                                Else
                                
                                ws.Cells(TickerCount, 11).Interior.ColorIndex = 3
                                
                                End If
                            
                        'Calculate and insert the total stock volume values
                        
                        ws.Cells(TickerCount, 12).Value = WorksheetFunction.Sum(ws.Cells(i, 7))
                        
        
            TickerCount = TickerCount + 1
            
           
            
            End If
            
            
        Next i
                            
                            
        'Before we do the summary, we are going to have to find the last row in the Tickercoloumn as we do not know how many sets we have in total
        
        LastrowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        
        'Now we need to get our summary values
        
              'Set up the table for Greatest increase, Greatest Decrease and Greatest total volume - using image provided in assignment instructions to pick where I place table
                ws.Cells(2, 15).Value = "Greatest % increase"
                ws.Cells(3, 15).Value = "Greatest % decrease"
                ws.Cells(4, 15).Value = "Greatest Total Volume"
                
                'Set up Ticker and Value
                
                ws.Cells(1, 16).Value = "Ticker"
                ws.Cells(1, 17).Value = "Value"
                
                
                'Here I am just telling it where to start looking from for these value
                
                Greatestvolume = ws.Cells(2, 12).Value
                Greatestincrease = ws.Cells(2, 11).Value
                Greatestdecrease = ws.Cells(2, 11).Value
                
                
                    'Now we can set up our next loop
                    
                        For i = 2 To LastrowI
                        
                        
                        'Looking for total greatest volume
                            If ws.Cells(i, 12).Value > Greatestvolume Then
                            Greatestvolume = ws.Cells(i, 12).Value
                            ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                            
                            
                        
                            Else
                            
                            Greatestvolume = Greatestvolume
                            
                            End If
                            
                            ws.Cells(4, 17).Value = Format(Greatestvolume, "Scientific")
                            
                        'Looking for greatest increase
                            If ws.Cells(i, 11).Value > Greatestincrease Then
                            Greatestincrease = ws.Cells(i, 11).Value
                            ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                            
                            Else
                            
                            Greatestincrease = Greatestincrease
                            
                            End If
                            
                            ws.Cells(2, 17).Value = Format(Greatestincrease, "Percent")
                            
                        
                        'Looking for greatest decrease
                        
                        If ws.Cells(i, 11).Value < Greatestdecrease Then
                            Greatestdecrease = ws.Cells(i, 11).Value
                            ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                            
                            Else
                            
                            Greatestdecrease = Greatestdecrease
                            
                            End If
                            
                          ws.Cells(3, 17).Value = Format(Greatestdecrease, "Percent")
                          
                          
                    Next i
                    
          'Go to next worksheet
            
            Next ws
                        
    

End Sub



