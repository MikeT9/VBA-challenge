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
        'define lastrow
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Add headers
        'add ticker to I1 and P1
        ws.Range("I1").Value = "Ticker"
        'add Yearly change to J1
        ws.Range("J1").Value = "Yearly Change"
        'add Percent Change to K1
        ws.Range("K1").Value = "Percent Change"
        'add Total Stock Volume to L1
        ws.Range("L1").Value = "Total Stock Volume"
        
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
              
    
    'end loop
    Next ws
    

End Sub
