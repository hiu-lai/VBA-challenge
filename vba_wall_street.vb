Sub SummarizeStock()
    ' Declare Variables
    Dim tickerCode As String
    Dim RowNum As Integer
    Dim CloseVal As Double
    Dim openVal As Double

    Dim PerChangeRange As Range
    Dim TotStkVol As Range
    
        
    TotSheets = Worksheets.Count
    
    For wksheets = 1 To TotSheets
        Worksheets(wksheets).Activate
        
         ' Count the number of rows in data table
        tblastRow = Cells(Rows.Count, 1).End(xlUp).Row
        
        'Clear Output areas
        Range("H1:Z" & tblastRow).ClearContents
        
        ' Count the number of columns in data table
        lastCol = Cells(1, Columns.Count).End(xlToLeft).Column
             
        ' Leaving two columns spacing from data table for data consolidation
        newOutputPos = lastCol + 3
        
        ' Set Heading for output table
        Cells(1, newOutputPos) = "Ticker"
        Cells(1, newOutputPos + 1) = "Yearly Change"
        Cells(1, newOutputPos + 2) = "Percent Change"
        Cells(1, newOutputPos + 3) = "Total Stock Volume"
        
        Cells(2, newOutputPos + 6) = "Greatest % Increase"
        Cells(3, newOutputPos + 6) = "Greatest % Decrease"
        Cells(4, newOutputPos + 6) = "Greatest Total Volume"
        
        Cells(1, newOutputPos + 7) = "Ticker"
        Cells(1, newOutputPos + 8) = "Value"
    
    
        'Column number for output table
        TickerCol = newOutputPos
        yearlyChangeCol = newOutputPos + 1
        PercentChangeCol = newOutputPos + 2
        StkValCol = newOutputPos + 3
        
        tickerOutput = newOutputPos + 7
        ValueOutput = newOutputPos + 8
        
        ' Set named new output data range
        Set PerChangeRangeL = Range(Cells(2, PercentChangeCol), Cells(tblastRow, PercentChangeCol))
        Set TotStkValM = Range(Cells(2, StkValCol), Cells(tblastRow, StkValCol))
        
   
        ' Set loop starting row
        StartRow = 2
        StockVal = Cells(StartRow, StkValCol).Value
    
        ' Set output start row num
        j = 1
        
        ' Loop for ticker code change
        For i = StartRow To tblastRow
            openVal = Cells(StartRow, 3).Value
            tickerCode = Cells(i, 1).Value
            ' Check for changes in ticker code then print and calculate output
            If (Cells(i, 1).Value = Cells(i + 1, 1).Value) Then
                StockVal = StockVal + Cells(i + 1, 7).Value
            Else:
                StartRow = i + 1
                j = j + 1
                CloseVal = Cells(i, 6).Value
                Cells(j, TickerCol).Value = Cells(i, 1).Value
                Cells(j, yearlyChangeCol).Value = CloseVal - openVal
                If openVal > 0 Then
                    Cells(j, PercentChangeCol).Value = Format(((CloseVal - openVal) / openVal), "Percent")
                End If
                Cells(j, StkValCol).Value = StockVal
                StockVal = Cells(i + 1, 7).Value
            End If
        Next i
        TotStkValM.NumberFormat = "0"
        Cells.EntireColumn.AutoFit
        
        ' Conditional Formatting
        '
        Dim ConditionalRange As Range
        Set ConditionalRange = PerChangeRangeL
        
        'Remove existing Conditions
        Cells.FormatConditions.Delete
        
        ' Condition 1 - Red if less than zero
        ConditionalRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
        ConditionalRange.FormatConditions(1).Interior.Color = RGB(255, 0, 0)
         
        ' Condition 2 - Green if greater than zero
        ConditionalRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
        ConditionalRange.FormatConditions(2).Interior.Color = RGB(0, 255, 0)
        
        'Bonus
        ' Populating Ticker Values
        Cells(2, ValueOutput).Value = Format(WorksheetFunction.Max(ConditionalRange), "Percent")
        Cells(3, ValueOutput).Value = Format(WorksheetFunction.Min(ConditionalRange), "Percent")
        Cells(4, ValueOutput).Value = WorksheetFunction.Max(TotStkValM)
        Cells(4, ValueOutput).NumberFormat = "0"
        
        Cells(2, tickerOutput) = WorksheetFunction.Index(Range("J2:J" & tblastRow), WorksheetFunction.Match(Cells(2, ValueOutput), ConditionalRange, 0))
        Cells(3, tickerOutput) = WorksheetFunction.Index(Range("J2:J" & tblastRow), WorksheetFunction.Match(Cells(3, ValueOutput), ConditionalRange, 0))
        Cells(4, tickerOutput) = WorksheetFunction.Index(Range("J2:J" & tblastRow), WorksheetFunction.Match(Cells(4, ValueOutput), TotStkValM, 0))
     Next wksheets
End Sub