Attribute VB_Name = "Module2"
Sub StockMarket():

'Steps:
'-------------

'Part I:
'1. Add column headers to begin a summary table with info about stocks for each year
'2. Loop through an excel file with 3 ws (each for a different year)
'3. For each ws display in the summary table the ticker symbol, yearly change, percent change, total stock volume
'4. Add conditional formatting for positive changes in green and negative changes in red

'To test scripts use excel file alpha-testing containing partial 2016 data
'Once scripts are functional, run on Stock-data file
'Stock-data is 3 ws (2014, 2015, 2016)each ws with A-Z ticker symbols for the year
    
'Part II:
'In a separate script
'Add an additional summary table that identifies the greatest % increase, % decrease and greatest total volume


    'Part I -
    '--------
           
    'Create variables and set counts to 0
    Dim Ticker As String
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim StockVolume As Double
    
    YearlyChange = 0
    PercentChange = 0
    StockVolume = 0
    
    'Determine number of the last row and last column
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    LastColumn = Cells(1, Columns.Count).End(xlToLeft).Column
    
    'Identify where to place a summary table and create the headers
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    'test "lasts"
    'MsgBox (LastRow)
    'MsgBox (LastColumn)

    'Id the open value for the first stock ticker, for calc of YearlyChange
    Dim OpenValue As Double
    OpenValue = Cells(2, 3).Value
    
    MsgBox (OpenValue)
    
    'Still to add, loop through each worksheet to create a summary table
    
    'Declare variables and create the for loop
    'Dim i As String
    For i = 2 To LastRow
                        
        'In column 1 look for changes in value when moving to next ticker symbol
        'If statement to identify what happens at the change in value for column 1
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
    
            'Add each different ticker symbol to the summary table during the loop
            Ticker = Cells(i, 1).Value
            Range("I" & Summary_Table_Row).Value = Ticker
            
            'Add Total stock volume to the summary table
            StockVolume = StockVolume + Cells(i, 7).Value
            Range("L" & Summary_Table_Row).Value = StockVolume
            
            'Add Yearly change for the stock to the summary table
            YearlyChange = Cells(i, 6).Value - OpenValue
            Range("J" & Summary_Table_Row).Value = YearlyChange
    
            'MsgBox (YearlyChange)
            
            'Add Percent change for the stock to the summary table
                   
            'Move down to next row of summary table for next loop
            Summary_Table_Row = Summary_Table_Row + 1
            
            'Reset the counter for StockVolume
            StockVolume = 0
            YearlyChange = 0
            'PercentChange = 0
                    
            'identify the next OpenValue
            OpenValue = Cells((i + 1), 3).Value
            'MsgBox (OpenValue)
                        
        Else
        
            StockVolume = StockVolume + Cells(i, 7).Value
                                      
        End If
        
    Next i
End Sub

