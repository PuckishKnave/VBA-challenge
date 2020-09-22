Attribute VB_Name = "Stock_Analysis"
Sub Stock_Analysis()
    
    'loop through all sheets
    For Each WS In Worksheets
        
        'Set variables
        Dim ticker As String
        Dim ticker_count As Long
        Dim total As Double
        Dim year_open As Double
        Dim year_close As Double
        Dim yearly_change As Double
        Dim percent_change As Double
        Dim lastrow As Double
        
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
        
        'Last row code column 1
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Cycle through the rows in column A
        For i = 2 To lastrow

            'Check for new ticker names
            If WS.Cells(i, 1).Value <> ticker Then

                'increase ticker count
                ticker_count = ticker_count + 1

                'Set new Ticker name
                ticker = WS.Cells(i, 1).Value

                'Set Year Open values
                year_open = WS.Cells(i, 3).Value

                'outputting row values for "Ticker" column
                WS.Cells(ticker_count, 9).Value = ticker

                'Set first total volume for new ticker
                total = WS.Cells(i, 7).Value
                WS.Cells(ticker_count, 12) = total

            Else
                'If ticker symbol is the same add to total
                total = total + WS.Cells(i, 7).Value
                WS.Cells(ticker_count, 12).Value = total
            End If
        
            'If it's the last of the specific ticker symbol
            If WS.Cells((i + 1), 1).Value <> ticker Then

                'output year close
                year_close = WS.Cells(i, 6).Value

                'Calculate yearly change
                yearly_change = year_close - year_open

                'Output to yearly change column
                WS.Cells(ticker_count, 10).Value = yearly_change

                'Calculate percent change
                If year_open = 0 Then
                    percent_change = yearly_change
                Else
                    percent_change = yearly_change / year_open
                End If

                'Output to percent change column
                WS.Cells(ticker_count, 11).Value = percent_change
                'Format to percent
                WS.Cells(ticker_count, 11).NumberFormat = "0.00%"

                'Color code positive and negative change
                If yearly_change > 0 Then
                    WS.Cells(ticker_count, 10).Interior.ColorIndex = 4
                ElseIf yearly_change < 0 Then
                    WS.Cells(ticker_count, 10).Interior.ColorIndex = 3
                
                End If
            
            End If
        Next i 
    Next WS
        
End Sub

