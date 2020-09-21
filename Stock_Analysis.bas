Attribute VB_Name = "Stock_Analysis"
Sub Stock_Analysis()
    
    'loop through all sheets
    For Each WS In Worksheets
        
        'Set variables
        Dim ticker As String
        Dim ticker_count As Double
        Dim total As Double
        Dim year_open As Double
        Dim year_close As Double
        
        ticker = ""
        ticker_count = 1
        total = 0
        year_open = 0
        year_close = 0
        
        'Create and name new columns
        WS.Range("I1").Value = "Ticker"
        WS.Range("J1").Value = "Yearly Change"
        WS.Range("K1").Value = "Percent Change"
        WS.Range("L1").Value = "Total Stock Volume"
        
        'Last row code
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        
        'Cycle through the rows in column A
        For Row = 2 To lastrow
        
    
End Sub

