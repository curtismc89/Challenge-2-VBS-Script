Attribute VB_Name = "Module1"
Sub stockAnalysis():

    Dim total As Double
    Dim row As Long
    Dim rowCount As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim summaryTableRow As Long
    Dim stockStartRow As Long
    Dim lastTicker As String
    
    For Each ws In Worksheets
    
     ws.Range("I1").Value = "Ticker"
     ws.Range("J1").Value = "Quarterly Change"
     ws.Range("K1").Value = "Percent Change"
     ws.Range("L1").Value = "Total Stock Volume"
     
     ws.Range("P1").Value = "Ticker"
     ws.Range("Q1").Value = "Value"
     ws.Range("O2").Value = "Greatest % Increase"
     ws.Range("O3").Value = "Greatest % Decrease"
     ws.Range("O4").Value = "Greatest Total Volume"
     
     summaryTableRow = 0
     total = 0
     quarterlyChange = 0
     stockStartRow = 2
     StartValue = 2
     
     rowCount = Cells(Rows.Count, "A").End(xlUp).row
     lastTicker = Cells(rowCount, 1).Value
     
     For row = 2 To rowCount
         If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then
             total = total + ws.Cells(row, 7).Value
             If total = 0 Then
                 ws.Range("I" & 2 + summaryTableRow).Value = ws.Cells(row, 1).Value
                 ws.Range("J" & 2 + summaryTableRow).Value = 0
                 ws.Range("K" & 2 + summaryTableRow).Value = 0
                 ws.Range("L" & 2 + summaryTableRow).Value = 0
             
             Else
                 If ws.Cells(StartValue, 3).Value = 0 Then
                     For findValue = StartValue To row
                         If ws.Cells(findValue, 3).Value <> 0 Then
                             StartValue = findValue
                             Exit For
                         End If
                     Next findValue
                 End If
                 quarterlyChange = ws.Cells(row, 6).Value - ws.Cells(StartValue, 3).Value
                 percentChange = quarterlyChange / ws.Cells(StartValue, 3).Value
                 ws.Range("I" & 2 + summaryTableRow).Value = ws.Cells(row, 1).Value
                 ws.Range("J" & 2 + summaryTableRow).Value = quarterlyChange
                 ws.Range("K" & 2 + summaryTableRow).Value = percentChange
                 ws.Range("L" & 2 + summaryTableRow).Value = total
                 If quarterlyChange > 0 Then
                     ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 4
                 ElseIf quarterlyChange < 0 Then
                     ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 3
                 Else
                     ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 0
                 End If
                 total = 0
                 averageChange = 0
                 quarterlyChange = 0
                 StartValue = row + 1
                 summaryTableRow = summaryTableRow + 1
                 
                                     
             End If
             
         Else
             total = total + ws.Cells(row, 7).Value
         End If
             
     Next row
             
     summaryTableRow = ws.Cells(Rows.Count, "I").End(xlUp).row
     Dim lastExtraRow As Long
     lastExtraRow = ws.Cells(Rows.Count, "J").End(xlUp).row
     For e = summaryTableRow To lastExtraRow
         For Column = 9 To 12
             ws.Cells(e, Column).Value = ""
             ws.Cells(e, Column).Interior.ColorIndex = 0
         Next Column
     Next e
             
     ws.Range("Q2").Value = WorksheetFunction.Max(ws.Range("K2:K" & summaryTableRow + 2))
     ws.Range("Q3").Value = WorksheetFunction.Min(ws.Range("K2:K" & summaryTableRow + 2))
     ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L2:L" & summaryTableRow + 2))
     
     Dim greatestIncreaseRow As Double
     Dim greatestDecreaseRow As Double
     Dim greatestTotVolRow As Double
     
     greatestIncreaseRow = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & summaryTableRow + 2)), ws.Range("K2:K" & summaryTableRow + 2), 0)
     greatestDecreaseRow = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & summaryTableRow + 2)), ws.Range("K2:K" & summaryTableRow + 2), 0)
     greatestTotVolRow = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & summaryTableRow + 2)), ws.Range("L2:L" & summaryTableRow + 2), 0)
     
     ws.Range("P2").Value = ws.Cells(greatestIncreaseRow + 1, 9).Value
     ws.Range("P3").Value = ws.Cells(greatestDecreaseRow + 1, 9).Value
     ws.Range("P4").Value = ws.Cells(greatestTotVolRow + 1, 9).Value
     
     For s = 0 To summaryTableRow
         ws.Range("J" & 2 + s).NumberFormat = "0.00"
         ws.Range("K" & 2 + s).NumberFormat = "0.00%"
         ws.Range("L" & 2 + s).NumberFormat = "#,###"
         
     Next s
     
     ws.Range("Q2").NumberFormat = "0.00%"
     ws.Range("Q3").NumberFormat = "0.00%"
     ws.Range("Q4").NumberFormat = "#,###"
     
     Columns("A:Q").AutoFit
   
   Next ws
    
End Sub



