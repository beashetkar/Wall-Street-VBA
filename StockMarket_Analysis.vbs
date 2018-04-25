'- PART 1
'-
'- Create a script that will loop through each year of stock data and grab the total amount of volume each stock had over the year.
'- You will also need to display the ticker symbol to coincide with the total volume.
'-
'- PART II
'-
'- Add the following info to Part I:
'- Yearly change from what the stock opened the year at to what the closing price was.
'- The percent change from the what it opened the year at to what it closed.
'- You should also have conditional formatting that will highlight positive change in green and negative change in red.
'-
'- PART III
'-
'- Your solution will also be able to locate the stock with the "Greatest % increase", "Greatest % Decrease" and
'- "Greatest total volume".
'-
'- CHALLENGE
'-
'- Make the appropriate adjustments to your script that will allow it to run on every worksheet just by running it once.


Sub StockMarket_Analysis()

    For Each ws In Worksheets
    
             
        '- Get Last Column and Last Row
            
        lcol = ws.Cells(1, Columns.Count).End(xlToLeft).Column
        lrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Add the words Ticker, Yearly_Change, Percentage_Change, Total_Volume to the Columns Headers
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly_Change"
        ws.Range("K1").Value = "Percentage_Change"
        ws.Range("L1").Value = "Total_Volume"
                
        'Add  the words Ticker & Value to the Column Headers For PART III
        
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        '- Declaration of variables
        
        '- For PART I
        
        Dim ticker As String
        Dim total_volume As Double
        
        '- For PART II
        
        Dim first_close, first_open As Double
        Dim yearly_percentage As Double
        Dim print_row As Integer

                
        '- For PART III
        
        Dim ticker_increase As String
        Dim ticker_decrease As String
        Dim ticker_volume As String
        
        Dim greatest_increase As Double
        Dim greatest_decrease As Double
        Dim greatest_volume As Double
        
        print_row = 2
        first_close = ws.Range("F2").Value
        first_open = ws.Range("C2").Value
        greatest_volume = ws.Cells(2, lcol).Value
        greatest_increase = 0
        greatest_decrease = 0
        total_volume = 0
                
        For Row = 2 To lrow
        
            If (ws.Cells(Row + 1, 1).Value <> ws.Cells(Row, 1).Value) And (first_close <> 0) Then
            
                '- For PART I : Add the Total Volume
                
                total_volume = total_volume + ws.Cells(Row, lcol).Value
                
                '- Print the Ticker and the Total Volume.
                
                ws.Range("I" & print_row).Value = ws.Cells(Row, 1).Value
                ws.Range("L" & print_row).Value = total_volume
                     
                '- For PART II : Calculate the yearly change and yearly percentage change and print it out
                
                yearly_percentage = (ws.Cells(Row, 6).Value - first_close) / first_close
                ws.Range("J" & print_row).Value = (ws.Cells(Row, 6).Value - first_close)
                ws.Range("K" & print_row).Value = FormatPercent(yearly_percentage)
                
                '- IF yearly change is negative set the cell color to red if it is positive, to green.
              
                If (ws.Cells(print_row, lcol + 3).Value < 0) Then
                    ws.Range("J" & print_row).Interior.ColorIndex = 3
                Else
                    ws.Range("J" & print_row).Interior.ColorIndex = 4
                End If
                
                '- For PART III : Compare greatest percentage increase, greatest percentage decrease, greatest_total volume
                
                If total_volume > greatest_volume Then
                
                    greatest_volume = total_volume
                    ticker_volume = ws.Cells(Row, 1).Value
                    
                End If
                
                If yearly_percentage > greatest_increase Then
                
                    greatest_increase = yearly_percentage
                    ticker_increase = ws.Cells(Row, 1).Value
                    
                ElseIf yearly_percentage < greatest_decrease Then
                
                    greatest_decrease = yearly_percentage
                    ticker_decrease = ws.Cells(Row, 1).Value
                    
                End If
                                
                ' Reset the Total_volume and increase print_row and get  the first_close of next ticker
               
                total_volume = 0
                print_row = print_row + 1
                first_close = ws.Cells(Row + 1, 6)
                
            ' If the cell immediately following a row is the same ticker...
            Else
                
                ' Add to the Total_Volume
                total_volume = total_volume + ws.Cells(Row, lcol).Value
            
            End If
       
        Next Row
        
        '- For Part III: Print Out greatest percentage increase, greatest percentage decrease, greatest_total volume
        
        ws.Range("P2").Value = ticker_increase
        ws.Range("Q2").Value = FormatPercent(greatest_increase)
        ws.Range("P3").Value = ticker_decrease
        ws.Range("Q3").Value = FormatPercent(greatest_decrease)
        ws.Range("P4").Value = ticker_volume
        ws.Range("Q4").Value = greatest_volume
        
 
    Next ws

End Sub



