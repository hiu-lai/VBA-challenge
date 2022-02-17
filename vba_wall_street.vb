Sub SummarizeStock()
    ' Declare Variables
    Dim tickerCode As String
    Dim RowNum As Integer
    Dim CloseVal As Double
    Dim openVal As Double

    Dim PerChangeRange As Range
    Dim TotStkVol As Range
    
    For Each ws In Worksheets
        
        ' Count the number of rows in data table
        tblastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Clear Output areas
        ws.Range("H1:ZZ" & tblastRow).ClearContents
        
        ' Count the number of columns in data table
        lastCol = ws.Cells(1, Columns.Count).End(xlToLeft).Column
             
        ' Leaving two columns spacing from data table for data consolidation
        newOutputPos = lastCol + 3
        
        ' Set Heading for output table
        ws.Cells(1, newOutputPos) = "Ticker"
        ws.Cells(1, newOutputPos + 1) = "Yearly Change"
        ws.Cells(1, newOutputPos + 2) = "Percent Change"
        ws.Cells(1, newOutputPos + 3) = "Total Stock Volume"
        
        ws.Cells(2, newOutputPos + 6) = "Greatest % Increase"
        ws.Cells(3, newOutputPos + 6) = "Greatest % Decrease"
        ws.Cells(4, newOutputPos + 6) = "Greatest Total Volume"
        
        ws.Cells(1, newOutputPos + 7) = "Ticker"
        ws.Cells(1, newOutputPos + 8) = "Value"
    
    
        'Column number for output table
        TickerCol = newOutputPos
        yearlyChangeCol = newOutputPos + 1
        PercentChangeCol = newOutputPos + 2
        StkValCol = newOutputPos + 3
        
        tickerOutput = newOutputPos + 7
        ValueOutput = newOutputPos + 8
        
        ' Set named new output data range
        Set PerChangeRangeL = ws.Range(ws.Cells(2, PercentChangeCol), ws.Cells(tblastRow, PercentChangeCol))
        Set TotStkValM = ws.Range(ws.Cells(2, StkValCol), ws.Cells(tblastRow, StkValCol))
        
   
        ' Set loop starting row
        StartRow = 2
        StockVal = ws.Cells(StartRow, StkValCol).Value
    
        ' Set output start row num
        j = 1
        
        ' Loop for ticker code change
        For i = StartRow To tblastRow
            openVal = ws.Cells(StartRow, 3).Value
            tickerCode = ws.Cells(i, 1).Value
            ' Check for changes in ticker code then print and calculate output
            If (ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value) Then
                StockVal = StockVal + ws.Cells(i + 1, 7).Value
            Else:
                StartRow = i + 1
                j = j + 1
                CloseVal = ws.Cells(i, 6).Value
                ws.Cells(j, TickerCol).Value = ws.Cells(i, 1).Value
                ws.Cells(j, yearlyChangeCol).Value = CloseVal - openVal
                If openVal > 0 Then
                    ws.Cells(j, PercentChangeCol).Value = Format(((CloseVal - openVal) / openVal), "Percent")
                End If
                ws.Cells(j, StkValCol).Value = StockVal
                StockVal = ws.Cells(i + 1, 7).Value
            End If
        Next i
        TotStkValM.NumberFormat = "0"
        ws.Cells.EntireColumn.AutoFit
        
        ' Conditional Formatting
        '
        Dim ConditionalRange As Range
        Set ConditionalRange = PerChangeRangeL
        
        'Remove existing Conditions
        ws.Cells.FormatConditions.Delete
        
        ' Condition 1 - Red if less than zero
        ConditionalRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
        ConditionalRange.FormatConditions(1).Interior.Color = RGB(255, 0, 0)
         
        ' Condition 2 - Green if greater than zero
        ConditionalRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
        ConditionalRange.FormatConditions(2).Interior.Color = RGB(0, 255, 0)
        
        'Bonus
        ' Populating Ticker Values
        ws.Cells(2, ValueOutput).Value = Format(WorksheetFunction.Max(ConditionalRange), "Percent")
        ws.Cells(3, ValueOutput).Value = Format(WorksheetFunction.Min(ConditionalRange), "Percent")
        ws.Cells(4, ValueOutput).Value = WorksheetFunction.Max(TotStkValM)
        ws.Cells(4, ValueOutput).NumberFormat = "0"
        
        ws.Cells(2, tickerOutput) = WorksheetFunction.Index(ws.Range("J2:J" & tblastRow), WorksheetFunction.Match(ws.Cells(2, ValueOutput), ConditionalRange, 0))
        ws.Cells(3, tickerOutput) = WorksheetFunction.Index(ws.Range("J2:J" & tblastRow), WorksheetFunction.Match(ws.Cells(3, ValueOutput), ConditionalRange, 0))
        ws.Cells(4, tickerOutput) = WorksheetFunction.Index(ws.Range("J2:J" & tblastRow), WorksheetFunction.Match(ws.Cells(4, ValueOutput), TotStkValM, 0))
     Next ws
End Sub


