Sub testing()
    For Each ws In Worksheets
        
        Dim Ticker_Name As String
        Cells(1, 9).Value = "Ticker"
        
        Dim Ticker_Total As Double
        Ticker_Total = 0
        
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        Cells(1, 10).Value = "Total Stock Volume"
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        For i = 2 To LastRow
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                Ticker_Name = Cells(i, 1).Value
                Ticker_Total = Ticker_Total + Cells(i, 7).Value
                
                Range("I" & Summary_Table_Row).Value = Ticker_Name
                Range("J" & Summary_Table_Row).Value = Ticker_Total

                Summary_Table_Row = Summary_Table_Row + 1

                Ticker_Total = 0
            Else
                Ticker_Total = Ticker_Total + Cells(i, 7).Value
            End If
        Next i
    Next ws
End Sub


