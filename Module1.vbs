Attribute VB_Name = "Module1"
Sub stock()

'create a loop through the worksheets for headers
    For Each ws In Worksheets
                
        'define variables
        'define Ticker
        Dim Ticker As String
        'define Row counter
        Dim Row_Counter As Double
        Row_Counter = 0
        'Set an initial variable for holding the yearly change per ticker
        Dim Yearly_Change As Double
        Yearly_Change = 0
        'Set an initial variable for holding the yearly Open per ticker
        Dim Yearly_Open As Double
        Yearly_Open = 0
        'Set an initial variable for holding the yearly Close per ticker
        Dim Yearly_Close As Double
        Yearly_Close = 0
        'Set an initial variable for holding the percent change per ticker
        Dim Perc_Change As Double
        Perc_Change = 0
        'Set an initial variable for holding the total stock volume per ticker
        Dim Total_Stock_Volume As Double
        Total_Stock_Volume = 0
        'Set an initial variable for holding the greatest stock increase
        Dim Max_Inc As Double
        Max_Inc = 0
        'Set an initial variable for holding the greatest stock decrease
        Dim Max_Dec As Double
        Max_Dec = 0
        'Set an initial variable for holding the greatest stock volume
        Dim Max_Vol As Double
        Max_Vol = 0
        'define lastrow
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'define lastcolumn
        lastcolumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
        
        'Add headers
        'add ticker to I1 and P1
        ws.Range("I1").Value = "Ticker"
        'add Yearly change to J1
        ws.Range("J1").Value = "Yearly Change"
        'add Percent Change to K1
        ws.Range("K1").Value = "Percent Change"
        'add Total Stock Volume to L1
        ws.Range("L1").Value = "Total Stock Volume"
        
        'Adding bonus table, outside of for loop
        ws.Cells(2, lastcolumn + 8).Value = "Greatest % Increase"
        ws.Cells(3, lastcolumn + 8).Value = "Greatest % Decrease"
        ws.Cells(4, lastcolumn + 8).Value = "Greatest Total Volume"
        ws.Cells(1, lastcolumn + 9).Value = "Ticker"
        ws.Cells(1, lastcolumn + 10).Value = "Value"
        
        'create a for loop for each set of values
        
        'create a loop based on the Ticker
        For i = 2 To lastrow
            
            'get ticker value
            Ticker = ws.Cells(i, 1).Value
            
            'find the yearly open value
            If Yearly_Open = 0 Then
                Yearly_Open = ws.Cells(i, 3).Value
            End If
            
            'add the values of the stock volumes together
            Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
            
            'Check if we are still within the same ticker, if it is not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                '+1 to row counter to set it to a 1 to exclude header row
                Row_Counter = Row_Counter + 1
                'add ticker value set to the sheet
                ws.Cells(Row_Counter + 1, 9) = Ticker
                
                'find the yearly close value
                Yearly_Close = ws.Cells(i, 6).Value
                
                'calc yearly change value
                Yearly_Change = Yearly_Close - Yearly_Open
                
                'add yearly change to sheets
                ws.Cells(Row_Counter + 1, 10).Value = Yearly_Change
                
                'if for perc change dont calc if 0
                If Yearly_Open <> 0 Then
                    'calc percentage change
                    Perc_Change = (Yearly_Change / Yearly_Open)
                End If
                
                'add percentage change to sheets
                ws.Cells(Row_Counter + 1, 11).Value = Perc_Change
                'change to percent format
                ws.Cells(Row_Counter + 1, 11).Value = FormatPercent(ws.Cells(Row_Counter + 1, 11).Value)
                
                'create if for colored changes in percentage
                If Perc_Change > 0 Then
                    ws.Cells(Row_Counter + 1, 11).Interior.ColorIndex = 4
                ElseIf Perc_Change < 0 Then
                    ws.Cells(Row_Counter + 1, 11).Interior.ColorIndex = 3
                End If
                
                'add total stock volume to worksheets
                ws.Cells(Row_Counter + 1, 12).Value = Total_Stock_Volume
                
                'reset values
                Yearly_Open = 0
                Yearly_Change = 0
                Perc_Change = 0
                Total_Stock_Volume = 0
            End If
        
        'end loop
        Next i

        'get derived values
        Max_Inc = WorksheetFunction.Max(ws.Range("K:K"))
        Max_Dec = WorksheetFunction.Min(ws.Range("K:K"))
        Max_Vol = WorksheetFunction.Max(ws.Range("L:L"))
        'add derived values to worksheet
        ws.Cells(2, 17).Value = Max_Inc
        ws.Cells(3, 17).Value = Max_Dec
        ws.Cells(4, 17).Value = Max_Vol
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        ws.Range("Q4:Q10").NumberFormat = "0"

        For i = 2 To lastrow

            'get ticker value
            Ticker = ws.Cells(i, 9).Value
            
            Debug.Print Ticker

            If ws.Cells(i, 11).Value = Max_Inc Then
                ws.Cells(2, 16).Value = Ticker
                Debug.Print Ticker
                
            ElseIf ws.Cells(i, 11).Value = Max_Dec Then
                ws.Cells(3, 16).Value = Ticker
                
            ElseIf Cells(i, 12).Value = Max_Vol Then
                ws.Cells(4, 16).Value = Ticker
            End If

        Next i
    
    'end loop
    Next ws
    

End Sub
