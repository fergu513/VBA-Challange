Sub Multiple_year_stock_data()
    
    'declare worksheet
    Dim ws As Worksheet
    
    'loop thru worksheets
    For Each ws In Worksheets
    
    'store results table and headers as true/false
    Dim Results_Sheet As Boolean
    Need_Summary_Table_Header = True
    
    'creating all my variables and setting them to 0
    Dim ticker As String
    Dim close_value As Double
    close_value = 0
    Dim open_value As Double

    Dim yearly_change As Double
    yearly_change = 0
    Dim percent_change As Double
    percent_change = 0
    
    'make variables for bonus values
    Dim Bonus_Increase As Double
    Bonus_Increase = 0
    Dim Bonus_Decrease As Double
    Bonus_Decrease = 0
    Dim Greatest_Volume As Double
    Greatest_Volume = 0
    
    'make variables for tickers of bonus values
    Dim Bonus_Increase_Ticker As String
    Dim Bonus_Decrease_Ticker As String
    Dim Greatest_Volume_Ticker As String
    
    'make headers for results table
    If Need_Summary_Table_Header Then
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    'make headers for bonus values
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    Else
    Need_Summary_Table_Header = True
    End If
    
    'starting the volume at 0
    Dim volume As Double
    volume = 0
    
    'keep track of which row of the data i am in
    Dim summary_table_row As Integer
    summary_table_row = 2
    
    'open value for the first ticker
    open_value = ws.Cells(2, 3).Value
    
    'making this applicable to data ranges of all sizes
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'row loop
    For i = 2 To last_row
    
    'take note whn the ticker ID changes
    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
    
    'closing value is the last value in column F before beginning a new ticket
    close_value = ws.Cells(i, 6).Value
    
    'yearly change is final value minus the first value
    yearly_change = close_value - open_value
    
    'percent change is yearly change / open_value
    percent_change = (yearly_change / open_value) * 100
    
    'print the yearly change in the summary table
    ws.Range("J" & summary_table_row).Value = yearly_change
    
    'print the yearly change in the summary table
    ws.Range("K" & summary_table_row).Value = CStr(percent_change) & "%"
    
    'note the ticker ID
    ticker = ws.Cells(i, 1).Value
    
    'add the last value for volume and take note
    volume = volume + ws.Cells(i, 7).Value
    
    'print the ticket in summary table
    ws.Range("I" & summary_table_row).Value = ticker
    
    'print the total volume
    ws.Range("L" & summary_table_row).Value = volume
    
    'color code yearly change
    If (yearly_change > 0) Then
    ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
    ElseIf (yearly_change <= 0) Then
    ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
    End If
    
    'color code percent change
    If (percent_change > 0) Then
    ws.Range("K" & summary_table_row).Interior.ColorIndex = 4
    ElseIf (percent_change <= 0) Then
    ws.Range("K" & summary_table_row).Interior.ColorIndex = 3
    End If
    
    'move to the next row of the summary table
    summary_table_row = summary_table_row + 1
    
    'calculate bonus values
    Bonus_Increase = WorksheetFunction.Max(ws.Range("K:K"))
    Bonus_Decrease = WorksheetFunction.Min(ws.Range("K:K"))
    Greatest_Volume = WorksheetFunction.Max(ws.Range("L:L"))
    
    'place bonus values in summary table
    ws.Range("Q2").Value = (CStr(Bonus_Increase) & "%")
    ws.Range("Q3").Value = (CStr(Bonus_Decrease) & "%")
    ws.Range("P2").Value = Bonus_Increase_Ticker
    ws.Range("P3").Value = Bonus_Decrease_Ticker
    ws.Range("Q4").Value = Greatest_Volume
    ws.Range("P4").Value = Greatest_Volume_Ticker
    
    'reset the volume back to 0 for the next ticker
    volume = 0
    
    'change the open value for the new ticker
    open_value = ws.Cells(i + 1, 3)

    Else
    
    'keep adding the volume
    volume = volume + ws.Cells(i, 7).Value
    
    End If
    
    Next i
    
    Next ws
    
End Sub