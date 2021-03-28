Attribute VB_Name = "Module2"
Sub StockMarketPart2():

    'Steps:
    '-------------

    'Part I:
    '1. Add column headers for a summary table with stock info for each year
    '2. Display in summary table for each stock: the ticker symbol, the yearly change, percent change, and total stock volume
    '3. Add conditional formatting for positive yearly changes in green and negative changes in red

        'To test scripts use excel file alpha-testing containing partial 2016 data
        'Once scripts are functional, run on Stock-data file
        'Stock-data is 3 ws (2014, 2015, 2016) each ws with A-Z ticker symbols for the year
    
    'Part II:
    'In a separate script, StockMarketPart2(), also do bonus activities
    '1. Using code from the first script, now loop through an excel file with 3 ws (each for a different year)
    '2. Add an additional summary table that identifies the greatest % increase, % decrease and greatest total volume


    'Part II - adding to part one the For Each ws and the additional summary table
    '--------
           
    'Loop through to create stock summary table for each ws
    For Each ws In Worksheets
    
        'Create variables and set counts to 0
        Dim Ticker As String
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim StockVolume As Double
    
        YearlyChange = 0
        PercentChange = 0
        StockVolume = 0
    
        'Declare variables for and determine number of the last row and last column
        Dim LastRow As Long
        Dim LastColumn As Integer
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
    
        'Identify where to place summary table and create the column headers
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
    
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Range("I1:L1").Interior.ColorIndex = 15
    
        'Adjust Column Widths
        ws.Columns("I").ColumnWidth = 8
        ws.Columns("J:K").ColumnWidth = 14
        ws.Columns("L").ColumnWidth = 18
    
        'Id the open value for the first stock ticker, for calc of YearlyChange
        Dim OpenValue As Double
        OpenValue = ws.Cells(2, 3).Value
    
        'test "lasts" and first stock's OpenValue
        'MsgBox (LastRow)
        'MsgBox (LastColumn)
        'MsgBox (OpenValue)
    
        'Declare variables and create for loop to create stock summary table
        Dim i As Long
    
        For i = 2 To LastRow
                        
            'In column 1 look for changes in value when moving to next ticker symbol
            'If statement to identify what happens at the change in value for column 1
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
    
                'Add each different ticker symbol to the summary table during the loop
                Ticker = ws.Cells(i, 1).Value
                ws.Range("I" & Summary_Table_Row).Value = Ticker
            
                'Add Yearly change for the stock to the summary table
                YearlyChange = ws.Cells(i, 6).Value - OpenValue
                ws.Range("J" & Summary_Table_Row).Value = YearlyChange
                
                
                    'Add Percent change for the stock to the summary table, and code to avoid Div/0
                    If OpenValue <> 0 Then
                                
                        PercentChange = YearlyChange / OpenValue
                        ws.Range("K" & Summary_Table_Row).Value = Format(PercentChange, "#,##0.00%")
                
                    Else
                
                        PercentChange = 0
                        ws.Range("K" & Summary_Table_Row).Value = Format(PercentChange, "#,##0.00%")
                                                            
                    End If
                
                'Test YearlyChange & PercentChange
                'MsgBox (YearlyChange)
                'MsgBox (PercentChange)
            
                'Add Total stock volume to the summary table
                StockVolume = StockVolume + ws.Cells(i, 7).Value
                ws.Range("L" & Summary_Table_Row).Value = StockVolume
                       
                'Move down to next row of summary table for next loop
                Summary_Table_Row = Summary_Table_Row + 1
            
                'Reset the counter for StockVolume
                StockVolume = 0
                YearlyChange = 0
                PercentChange = 0
                    
                'identify/store the next stock's OpenValue for the next loop
                OpenValue = ws.Cells((i + 1), 3).Value
                
                'MsgBox (OpenValue)
                        
            Else
        
                StockVolume = StockVolume + ws.Cells(i, 7).Value
                                      
            End If
    
        Next i
    
        'Declare variable and determine number of the last row of the summary table
        Dim LastRowSummary As Long
        LastRowSummary = ws.Cells(Rows.Count, 10).End(xlUp).Row
    
        'Loop through summary table to fill color red for losses and green for gains
        For i = 2 To LastRowSummary
               
            If ws.Cells(i, 10).Value > 0 Then
         
                ws.Cells(i, 10).Interior.ColorIndex = 4
            
            ElseIf ws.Cells(i, 10).Value < 0 Then
        
                ws.Cells(i, 10).Interior.ColorIndex = 3
            
            Else
        
            'if there was no gain or loss
                ws.Cells(i, 10).Interior.ColorIndex = 0
            
            End If
          
        Next i
    
        
        'find the largest percent in column K and enter that value in Q2
        
        'Create variables for analysis table and set counts to 0
        Dim MaxPercent As Double
        Dim MinPercent As Double
        Dim MaxStockVolume As Double
        
        MaxPercent = 0
        MinPercent = 0
        MaxStockVolume = 0
        
        'Identify where to place analysis table and create the column and row labels for table
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        'Adjust Column Widths for analysis table
        ws.Columns("O").ColumnWidth = 18
        ws.Columns("Q").ColumnWidth = 18
                       
        'update the color of the header row for analysis table
        ws.Range("P1:Q1").Interior.ColorIndex = 15
              
        'Next section of code incomplete/not functioning
        'For i = 2 To LastRowSummary
        
            'If ws.Cells(i, 12).Value > ws.Cells(i + 1, 12).Value Then
                
              'MaxStockVolume = ws.Cells(i, 12).Value
                
            'Else
            
              'MaxStockVolume = ws.Cells(i + 1, 12).Value
           
            'End If
        
            'ws.Cells(3, 18).Value = MaxStockVolume
        
        'Next i
        
        
        'For i = 2 To LastRowSummary
        
            'If ws.Cells(i, 11).Value > ws.Cells(i + 1, 11).Value Then
                
                'MaxPercent = ws.Cells(i, 11).Value
                
            'Else
            
                'MaxPercent = ws.Cells(i + 1, 11).Value
           
            'End If
        
            'ws.Cells(2, 17).Value = MaxPercent
        
        'Next i
                  
        'For i = 2 To LastRowSummary
        
            'If ws.Cells(i, 11).Value < ws.Cells(i + 1, 11).Value Then
                
                'MinPercent = ws.Cells(i, 11).Value
                
            'Else
            
                'MinPercent = ws.Cells(i + 1, 11).Value
           
            'End If
        
            'ws.Cells(2, 17).Value = MinPercent
        
        'Next i
    Next ws
    
End Sub


