Sub Stocks()

    'Loop through all the sheets
    For Each ws In Worksheets
    
        'Declaring my variables
        'Initial variable for holding the ticker name
        Dim Ticker_Name As String
    
        'Variable for the total yearly change per ticker
        Dim Ticker_Yearly_Change As Double
        Dim Percent_Change As Double
        Dim Ticker_Open As Double
        Dim Ticker_Close As Double
    
        Dim Row_Counter As Integer
        Row_Counter = 2
    
        'Finds the first cell with the ticker open and assigns that value
        Ticker_Open = ws.Cells(Row_Counter, 3).Value
    
        Dim Total_Stock_Vol As Double
    
    
        'Creates headers for new columns
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
            
        'All of the row values in the column being searched
        For i = Row_Counter To 753001
            
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            
            Ticker_Name = ws.Cells(i, 1).Value
            Ticker_Close = ws.Cells(i, 6).Value
            Ticker_Yearly_Change = Ticker_Close - Ticker_Open
            Percent_Change = Ticker_Yearly_Change / Ticker_Open
            Total_Stock_Vol = Total_Stock_Vol + ws.Cells(i, 7).Value
            
            'Conditional formatting for when the ticker yearly change is >/< 0
            If Ticker_Yearly_Change > 0 Then
            ws.Range("J" & Row_Counter).Interior.ColorIndex = 4
            ElseIf Ticker_Yearly_Change < 0 Then
            ws.Range("J" & Row_Counter).Interior.ColorIndex = 3
            
            End If
            
            'Prints the values in the new columns
            ws.Range("I" & Row_Counter).Value = Ticker_Name
            ws.Range("J" & Row_Counter).Value = Ticker_Yearly_Change
            ws.Range("K" & Row_Counter).Value = FormatPercent(Percent_Change, 2)
            ws.Range("L" & Row_Counter).Value = Total_Stock_Vol
            
            'Sets a new value to ticker open once the ticker symbol is different
            Ticker_Open = Cells(i + 1, 3).Value
            
            'Adds 1 to the row value
            Row_Counter = Row_Counter + 1
            'Resets the total stock volume
            Total_Stock_Vol = 0
            
            Else
            
            Total_Stock_Vol = Total_Stock_Vol + ws.Cells(i, 7).Value
            
            End If
        
        Next i
        
        'Need to reset row counter back to 2
        Row_Counter = 2
        Lastrow_Summary_Table = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        For i = Row_Counter To Lastrow_Summary_Table
        
            If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & Lastrow_Summary_Table)) Then
            ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
            ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
            ws.Cells(2, 17).NumberFormat = "0.00%"
            
            ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & Lastrow_Summary_Table)) Then
            ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
            ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
            ws.Cells(3, 17).NumberFormat = "0.00%"
            
            ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & Lastrow_Summary_Table)) Then
            ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
            ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
            
            End If
            
        Next i
    
    Next ws
        
End Sub
