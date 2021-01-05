Attribute VB_Name = "Module1"
Sub StockAnalysis()
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    
    Columns("L").ColumnWidth = 20
    Columns("O").ColumnWidth = 20
    Columns("Q").ColumnWidth = 20
    
    Dim ticker As String
    Dim tickerIndex As Integer
    Dim openValue As Double
    Dim closeValue As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalStockVol As Double
    Dim lastRowIncDec As Long
    Dim lastRowTotalVol As Long
    tickerIndex = 2
    openValue = Cells(2, 3).Value
    ActiveSheet.range("K:K").NumberFormat = "0.00%"
    
    For Index = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        ticker = Cells(Index, 1).Value
        totalStockVol = totalStockVol + Cells(Index, 7).Value
        
        If ticker <> Cells(Index + 1, 1).Value Then
            
            Cells(tickerIndex, 9).Value = ticker
            closeValue = Cells(Index, 6).Value
            Cells(tickerIndex, 12).Value = totalStockVol
            yearlyChange = openValue - closeValue
            Cells(tickerIndex, 10).Value = yearlyChange
        If yearlyChange > 0 Then
            Cells(tickerIndex, 10).Interior.Color = vbGreen
            Else
            Cells(tickerIndex, 10).Interior.Color = vbRed
        End If
        If openValue > 0 Then
            ''percentChange = Format((yearlyChange / openValue) * 100, "0.00")
            ''Cells(tickerIndex, 11).Value = Str(percentChange) + "%"
            
            percentChange = yearlyChange / openValue
            Cells(tickerIndex, 11).Value = percentChange
        Else
            Cells(tickerIndex, 11).Value = "NA"
        End If
            openValue = Cells(Index + 1, 3).Value
            totalStockVol = 0
            tickerIndex = tickerIndex + 1
            
        End If
        
    Next Index
    
    With Application.ActiveSheet
        lastRowIncDec = .range("K" & .Rows.Count).End(xlUp).Row
    End With
    
    With Application.ActiveSheet
        lastRowTotalVol = .range("L" & .Rows.Count).End(xlUp).Row
    End With
    
    range("Q2").NumberFormat = "0.00%"
    range("Q3").NumberFormat = "0.00%"
    
    range("Q2").Value = WorksheetFunction.Max(range("K2:K" & lastRowIncDec))
    range("Q3").Value = WorksheetFunction.Min(range("K2:K" & lastRowIncDec))
    range("Q4").Value = WorksheetFunction.Max(range("L2:L" & lastRowTotalVol))
    
    Dim maxRowInc As Long
    Dim pctValue As String
    ''pctValue = Format(range("Q2").Value, "Percent")
    pctValue = CStr(range("Q2").Value * 100) + "%"
    maxRowInc = ActiveSheet.range("K:K").Find(What:=pctValue).Row
    range("P2").Value = Cells(maxRowInc, 9)
    
    Dim maxRowDec As Long
    Dim pctValue2 As String
    ''pctValue2 = Format(range("Q3").Value, "Percent")
    pctValue2 = CStr(range("Q3").Value * 100) + "%"
    maxRowDec = ActiveSheet.range("K:K").Find(What:=pctValue2).Row
    range("P3").Value = Cells(maxRowDec, 9)
    
    Dim maxRowTotVol As Long
    Dim Value3 As String
    Value3 = range("Q4").Value
    maxRowTotVol = ActiveSheet.range("L:L").Find(What:=Value3).Row
    range("P4").Value = Cells(maxRowTotVol, 9)
    
End Sub
