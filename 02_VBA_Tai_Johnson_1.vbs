Sub Stock_analysis()

        Dim ws As Worksheet
        
        For Each ws In Worksheets
        
        'Set up the variables needed to store data as we go
        'Being that calendar years aren't always the same number of business days
        'we'll set this up based on the each ticker symbol changing or not rather than by dates
        
        Dim open_val As Double
        Dim close_val As Double
        Dim volume As Double
        Dim ticker_sym As String
        Dim results_counter As Double
        Dim ticker_counter As Double
            
        'Variables to handle greatest increase, decrease and volume
        Dim grtst_inc As Double
        Dim grtst_decr As Double
        Dim grtst_vol As Double
        Dim grtst_counter As Long
        Dim grtst_inc_sym As String
        Dim grtst_dec_sym As String
        Dim grtst_vol_sym As String
        
        'Label our results columns
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly change"
        ws.Cells(1, 11).Value = "Percent change"
        ws.Cells(1, 12).Value = "Total Volume"
        
        'Start code at the first ticker symbol and set our "ticker counter"
        ticker_counter = 2
        results_counter = 2
        ticker_sym = ws.Cells(2, 1).Value
            
        'Outer loop iterating through the whole sheet
        Do Until IsEmpty(ws.Cells(ticker_counter, 1).Value) = True
            
            'Report current ticker symbol and log stock opening value (only log open value if it is greater than 0,
            'if it's 0, code will throw a "divide by 0" error.
            ws.Cells(results_counter, 9).Value = ticker_sym
            open_val = ws.Cells(ticker_counter, 3).Value
                    
            'Outer loop iterating through each particular ticker symbol
            Do Until ws.Cells(ticker_counter, 1).Value <> ticker_sym
                
                volume = volume + ws.Cells(ticker_counter, 7).Value
                close_val = ws.Cells(ticker_counter, 6).Value
                ticker_counter = ticker_counter + 1
                
            Loop
            
            'Log next new ticker symbol and report yearly change, percent change and total volume
            'change color of yearly change based on negative or postive
            ticker_sym = ws.Cells(ticker_counter, 1).Value
            ws.Cells(results_counter, 10).Value = close_val - open_val
            If ws.Cells(results_counter, 10).Value > 0 Then
                ws.Cells(results_counter, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(results_counter, 10).Interior.ColorIndex = 3
            End If
            
            'Handle cases where the stock does not hit the market until later into the year, in short
            'it opens the year at 0
            If open_val = 0 Then
                ws.Cells(results_counter, 11).Value = FormatPercent(1)
            Else
                ws.Cells(results_counter, 11).Value = FormatPercent((close_val - open_val) / open_val)
            End If
            
            ws.Cells(results_counter, 12).Value = volume
            results_counter = results_counter + 1
            volume = 0
                   
        Loop
        
        'Report greatest increase, decrease and volume
        
        'Iterate through our final results info to find greatest value of each category and it's
        'corresponding ticker
        'first record the first ticker in case it is the greatest in each category
        grtst_counter = 2
        If ws.Cells(grtst_counter, 11).Value > 0 Then grtst_inc = ws.Cells(grtst_counter, 11).Value
        If ws.Cells(grtst_counter, 11).Value < 0 Then grtst_decr = ws.Cells(grtst_counter, 11).Value
        grtst_vol = ws.Cells(grtst_counter, 12).Value
        
        grtst_inc_sym = ws.Cells(grtst_counter, 9).Value
        grtst_dec_sym = ws.Cells(grtst_counter, 9).Value
        grtst_vol_sym = ws.Cells(grtst_counter, 9).Value
        
        Do Until IsEmpty(ws.Cells(grtst_counter, 9).Value) = True
            
            If ws.Cells(grtst_counter, 11).Value > grtst_inc Then
                grtst_inc = ws.Cells(grtst_counter, 11).Value
                grtst_inc_sym = ws.Cells(grtst_counter, 9).Value
            End If
            
            If ws.Cells(grtst_counter, 11).Value < grtst_decr Then
                grtst_decr = ws.Cells(grtst_counter, 11).Value
                grtst_dec_sym = ws.Cells(grtst_counter, 9).Value
            End If
            
            If ws.Cells(grtst_counter, 12).Value > grtst_vol Then
                grtst_vol = ws.Cells(grtst_counter, 12).Value
                grtst_vol_sym = ws.Cells(grtst_counter, 9).Value
            End If
            
            grtst_counter = grtst_counter + 1
        
        Loop
        
        'Label the section for our "greastest" section and report our findings to our spreadsheet
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest total volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Greatest % Increase"
        
        ws.Cells(2, 15).Value = grtst_inc_sym
        ws.Cells(2, 16).Value = FormatPercent(grtst_inc)
        ws.Cells(3, 15).Value = grtst_dec_sym
        ws.Cells(3, 16).Value = FormatPercent(grtst_decr)
        ws.Cells(4, 15).Value = grtst_vol_sym
        ws.Cells(4, 16).Value = grtst_vol
        
        ws.Columns("I:P").AutoFit
        
    Next ws
        
End Sub