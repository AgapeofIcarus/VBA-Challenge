'To generate the ticker, yearly price change, percent change, and total stock volume

Sub StockData()
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate
        ' Determine the Last Row
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

        ' Add Column Headers
        Cells(1, "I").Value = "Ticker"
        Cells(1, "J").Value = "Yearly Change"
        Cells(1, "K").Value = "Percent Change"
        Cells(1, "L").Value = "Total Stock Volume"
        
        'Determine the Variables
        Dim OpeningPr As Double
        Dim ClosingPr As Double
        Dim YearlyCh As Double
        Dim Ticker As String
        Dim PercentCh As Double
        Dim Stock As Double
        Volume = 0
        Dim Row As Double
        Row = 2
        Dim Column As Integer
        Column = 1
        Dim i As Long
        
        'Set Starting point for Open Price Loop
        OpeningPr = Cells(2, Column + 2).Value
        
         ' Set the ticker loop to generate in designated cells
        For i = 2 To LastRow
        
         ' If the cells have differing tickers
            If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
            
                ' Set Ticker name
                Ticker = Cells(i, Column).Value
                Cells(Row, Column + 8).Value = Ticker
                
                ' Set Closing Price
                ClosingPr = Cells(i, Column + 5).Value
                
                ' Add Yearly Change
                YearlyCh = ClosingPr - OpeningPr
                Cells(Row, Column + 9).Value = YearlyCh
                
                ' Add Percent Change
                If (OpeningPr = 0 And ClosingPr = 0) Then
                    PercentCh = 0
                ElseIf (OpeningPr = 0 And ClosingPr <> 0) Then
                    PercentCh = 1
                Else
                    PercentCh = YearlyCh / OpeningPr
                    Cells(Row, Column + 10).Value = PercentCh
                    Cells(Row, Column + 10).NumberFormat = "0.00%"
                End If
                
                'Add Total Stock
                Stock = Stock + Cells(i, Column + 6).Value
                Cells(Row, Column + 11).Value = Volume
                
                ' Add one to the summary table row
                Row = Row + 1
                
                ' reset the Opening Price
                OpeningPr = Cells(i + 1, Column + 2)
                
                ' reset the Volumn Total
                Volume = 0
                
            'If the cells are the same ticker
            Else
                Volume = Volume + Cells(i, Column + 6).Value
            End If
        Next i
        
        ' Determine the Last Row of Yearly Change per WS
        YCLastRow = WS.Cells(Rows.Count, Column + 8).End(xlUp).Row
        
        ' Set the Cell Colors to change depending on positive or negative trends
        For j = 2 To YCLastRow
            If (Cells(j, Column + 9).Value > 0 Or Cells(j, Column + 9).Value = 0) Then
                Cells(j, Column + 9).Interior.ColorIndex = 10
            ElseIf Cells(j, Column + 9).Value < 0 Then
                Cells(j, Column + 9).Interior.ColorIndex = 3
            End If
        Next j
        
        ' Set Greatest % Increase, % Decrease, and Total Volume headers
        Cells(2, Column + 14).Value = "Greatest % Increase"
        Cells(3, Column + 14).Value = "Greatest % Decrease"
        Cells(4, Column + 14).Value = "Greatest Total Volume"
        Cells(1, Column + 15).Value = "Ticker"
        Cells(1, Column + 16).Value = "Value"
        
        ' Search WS to find the needed values for greatest % increase, decrease, and total volume
        For Z = 2 To YCLastRow
            If Cells(Z, Column + 10).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & YCLastRow)) Then
                Cells(2, Column + 15).Value = Cells(Z, Column + 8).Value
                Cells(2, Column + 16).Value = Cells(Z, Column + 10).Value
                Cells(2, Column + 16).NumberFormat = "0.00%"
            ElseIf Cells(Z, Column + 10).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & YCLastRow)) Then
                Cells(3, Column + 15).Value = Cells(Z, Column + 8).Value
                Cells(3, Column + 16).Value = Cells(Z, Column + 10).Value
                Cells(3, Column + 16).NumberFormat = "0.00%"
            ElseIf Cells(Z, Column + 11).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & YCLastRow)) Then
                Cells(4, Column + 15).Value = Cells(Z, Column + 8).Value
                Cells(4, Column + 16).Value = Cells(Z, Column + 11).Value
            End If
        Next Z
        
    Next WS
        
End Sub
