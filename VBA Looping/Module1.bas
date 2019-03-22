Attribute VB_Name = "Module1"
Sub StockVolume():

'Loop through workbook
For Each ws In Worksheets

    'Set variable to hold ticker symbols
    Dim Ticker As String
    'Set variable to hold Stock Total
    Dim Total_Stock_Volume As LongLong
    Total_Stock_Volume = 0
    'Create summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    'Find Last Row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    'Add Ticker to Column Header
    ws.Cells(1, 9).Value = "Ticker"
    'Add Total Stock Volume to Column Header
    ws.Cells(1, 10).Value = "Total Stock Volume"
        'Loop through worksheet
        For i = 2 To LastRow
                
        'Find changes in ticker symbol
        If (ws.Cells(i + 1, 1) <> ws.Cells(i, 1)) Then
        
        'Set Ticker
        Ticker = ws.Cells(i, 1).Value
        
        'Add Total Stock Volume
        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
        
        'Place unique list
        ws.Range("I" & Summary_Table_Row).Value = Ticker
        
        'Place Total_Stock_Volume
        ws.Range("J" & Summary_Table_Row).Value = Total_Stock_Volume
        
        'Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
        
        'Reset Total_Stock_Volume
        Total_Stock_Volume = 0
        
        'If the following cell contains the same ticker
        Else
        
        'Add to Total_Stock_Volume
        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                    
        End If
    
        Next i
        'Increase column width
        ws.Columns("A:J").AutoFit
Next ws
End Sub

