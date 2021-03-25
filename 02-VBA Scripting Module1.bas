Attribute VB_Name = "Module1"
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


    'Part 1
    '--------
           
    'Loop through all the stocks for one year
    'Add each different ticker symbol to the summary table
    'Add Yearly change for the stock to the summary table
    'Add Percent change for the stock to the summary table
    'Add Total stock volume to the summary table
    
        
    'Create variables and set counts to 0
    Dim Ticker As String
    Dim YearlyChange As Double
    Dim StockVolume As Double
    
    YearlyChange = 0
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
    MsgBox (LastRow)
    MsgBox (LastColumn)

    'need a way to calc the YearlyChange
    Dim OpenValue As Double
    OpenValue = Cells(2, 3).Value
    
    MsgBox (OpenValue)
    
    'Loop creation - will need to loop through each worksheet
    
    'Declare variables and create a for loop
                
    'In column 1 look for changes in value when moving to next ticker symbol
    'If statement to identify what happens at the change in value for column 1
    
End Sub
