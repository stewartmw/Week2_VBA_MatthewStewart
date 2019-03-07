Option Explicit

Dim LastRow As Long
Dim LastColumn As Integer
Dim SummaryTableRow As Integer
Dim SummaryTableColumn1, SummaryTableColumn2, SummaryTableColumn3, SummaryTableColumn4 As Integer
Dim SummaryTableLastRow As Integer
Dim SummaryTableFirstColumnLetter As String
Dim SummaryTableLastColumnLetter As String
Dim MaxSummaryTableColumn1, MaxSummaryTableColumn2, MaxSummaryTableColumn3 As Integer
Dim MaxSummaryTableFirstColumnLetter As String
Dim MaxSummaryTableLastColumnLetter As String
Dim MaxIncrease As Double
Dim MaxDecrease As Double
Dim MaxVolume As Variant
Dim YearOpenValue, YearCloseValue As Double
Dim TotalStockVolume As Variant
Dim i As Long
Dim j As Integer
Dim ws As Worksheet

Sub StockAnalysis():

    For Each ws In Worksheets

        'Identify last Row and Column numbers in initial data
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
    
        'Initial setup of summary table and final ("maximums") summary table
        '-------------------------------------------------------------------
        SummaryTableColumn1 = LastColumn + 2
        SummaryTableColumn2 = LastColumn + 3
        SummaryTableColumn3 = LastColumn + 4
        SummaryTableColumn4 = LastColumn + 5
    
        SummaryTableRow = 1 'Initialize SummaryTableRow for use in "i" For Loop below
    
        ws.Cells(1, SummaryTableColumn1).Value = "Ticker"
        ws.Cells(1, SummaryTableColumn2).Value = "Yearly Change"
        ws.Cells(1, SummaryTableColumn3).Value = "Percent Change"
        ws.Cells(1, SummaryTableColumn4).Value = "Total Stock Volume"
        
        MaxSummaryTableColumn1 = LastColumn + 7
        MaxSummaryTableColumn2 = LastColumn + 8
        MaxSummaryTableColumn3 = LastColumn + 9
        
        ws.Cells(2, MaxSummaryTableColumn1).Value = "Greatest % Increase"
        ws.Cells(3, MaxSummaryTableColumn1).Value = "Greatest % Decrease"
        ws.Cells(4, MaxSummaryTableColumn1).Value = "Greatest Total Volume"
        ws.Cells(1, MaxSummaryTableColumn2).Value = "Ticker"
        ws.Cells(1, MaxSummaryTableColumn3).Value = "Value"
        '-------------------------------------------------------------------
        
        'Generate summary table
        '-------------------------------------------------------------------
        For i = 2 To LastRow
        
            If (ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value) Then
                'Create new row for summary table; identify new ticker symbol;
                    'initialize opening stock value and opening stock volume
                SummaryTableRow = SummaryTableRow + 1
                ws.Cells(SummaryTableRow, SummaryTableColumn1).Value = ws.Cells(i, 1).Value
                YearOpenValue = ws.Cells(i, 3).Value
                TotalStockVolume = CDec(ws.Cells(i, 7).Value)
            Else
                'Aggregate stock volume by ticker symbol as for loop iterates;
                    'identify final closing value of each stock
                If YearOpenValue = 0 Then
                    YearOpenValue = ws.Cells(i, 3).Value
                End If
                TotalStockVolume = CDec(TotalStockVolume + ws.Cells(i, 7).Value)
                YearCloseValue = ws.Cells(i, 6).Value
            End If
        
            'Populate Columns 2-4 of summary table
            ws.Cells(SummaryTableRow, SummaryTableColumn2).Value = YearCloseValue - YearOpenValue
        
            If YearOpenValue = 0 Then
                'Protect against Divide By 0 Error
                ws.Cells(SummaryTableRow, SummaryTableColumn3).Value = 0
            Else
                ws.Cells(SummaryTableRow, SummaryTableColumn3).Value = _
                (YearCloseValue - YearOpenValue) / YearOpenValue
            End If
        
            ws.Cells(SummaryTableRow, SummaryTableColumn4).Value = TotalStockVolume
        
        Next i
        '--------------------------------------------------------------------------
    
        'Format summary table and generate final ("maximums") summary table
        '----------------------------------------------------------------------------
        SummaryTableLastRow = ws.Cells(Rows.Count, SummaryTableColumn1).End(xlUp).Row
    
        'Initialize MaxIncrease, MaxDecrease, and MaxVolume
        MaxIncrease = ws.Cells(2, SummaryTableColumn3).Value
        MaxDecrease = ws.Cells(2, SummaryTableColumn3).Value
        MaxVolume = CDec(ws.Cells(2, SummaryTableColumn4).Value)
        
        For j = 2 To SummaryTableLastRow
        
            If ws.Cells(j, SummaryTableColumn2).Value >= 0 Then
            'Green if yearly change >= 0; red otherwise
                ws.Cells(j, SummaryTableColumn2).Interior.ColorIndex = 4
            Else
                ws.Cells(j, SummaryTableColumn2).Interior.ColorIndex = 3
            End If
        
            'Format percent change as percentages, not decimals
            ws.Cells(j, SummaryTableColumn3) = _
                FormatPercent(ws.Cells(j, SummaryTableColumn3), 2, vbTrue)
            
        'Populate final ("maximums") summary table
            If ws.Cells(j, SummaryTableColumn3).Value > MaxIncrease Then
            'Identify max percent change and populate final table
                MaxIncrease = ws.Cells(j, SummaryTableColumn3).Value
                ws.Cells(2, MaxSummaryTableColumn2).Value = _
                    ws.Cells(j, SummaryTableColumn1).Value
                ws.Cells(2, MaxSummaryTableColumn3).Value = MaxIncrease
                ws.Cells(2, MaxSummaryTableColumn3) = _
                FormatPercent(ws.Cells(2, MaxSummaryTableColumn3), 2, vbTrue)
            End If
                
            If ws.Cells(j, SummaryTableColumn3).Value < MaxDecrease Then
            'Identify max (negative) percent change and populate final table
                MaxDecrease = ws.Cells(j, SummaryTableColumn3).Value
                ws.Cells(3, MaxSummaryTableColumn2).Value = _
                    ws.Cells(j, SummaryTableColumn1).Value
                ws.Cells(3, MaxSummaryTableColumn3).Value = MaxDecrease
                ws.Cells(3, MaxSummaryTableColumn3) = _
                FormatPercent(ws.Cells(3, MaxSummaryTableColumn3), 2, vbTrue)
            End If
            
            If CDec(ws.Cells(j, SummaryTableColumn4).Value) > MaxVolume Then
            'Identify max total volume and populate final table
                MaxVolume = CDec(ws.Cells(j, SummaryTableColumn4).Value)
                ws.Cells(4, MaxSummaryTableColumn2).Value = _
                    ws.Cells(j, SummaryTableColumn1).Value
                ws.Cells(4, MaxSummaryTableColumn3).Value = MaxVolume
            End If
            
        Next j
    
        'Autofit column widths in both summary tables; had to convert column numbers to letters
        SummaryTableFirstColumnLetter = Split(ws.Cells(1, SummaryTableColumn1).Address, "$")(1)
        SummaryTableLastColumnLetter = Split(ws.Cells(1, SummaryTableColumn4).Address, "$")(1)
    
        ws.Columns(SummaryTableFirstColumnLetter & ":" & SummaryTableLastColumnLetter).AutoFit
        
        MaxSummaryTableFirstColumnLetter = Split(ws.Cells(1, MaxSummaryTableColumn1).Address, "$")(1)
        MaxSummaryTableLastColumnLetter = Split(ws.Cells(1, MaxSummaryTableColumn3).Address, "$")(1)
    
        ws.Columns(MaxSummaryTableFirstColumnLetter & ":" & MaxSummaryTableLastColumnLetter).AutoFit
    
    Next ws
    
End Sub
