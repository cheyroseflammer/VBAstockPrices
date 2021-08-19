Attribute VB_Name = "stockScript"
Sub stockScript()
  ' Setting up variables
  Dim tickerSymbol As String
  Dim yearlyChange As Double
  Dim percentChange As Double
  Dim totalStockVol As LongLong
  Dim openPrice
  Dim closePrice As Double
  Dim summaryTableRow As Integer
  Dim yearStart As Integer
  Dim lastRow As Double
  '----------------------------------------------------------------
  ' Loop through all the sheets
  '----------------------------------------------------------------
  For Each ws In Worksheets
    ' Determine last row
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    ' Makes 2nd row equal where entries get filled in as loop moves
    summaryTableRow = 2
    ' Setting start val
    totalStockVol = 0
    ' Make table headers
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percentage Change"
    ws.Range("L1") = "Total Stock Volume"
    ' Temp open
    openPrice = ws.Cells(2, 3).Value
    ' Loop through the rows
    For i = 2 To lastRow
      tickerSymbol = ws.Cells(i, 1).Value
      If tickerSymbol <> ws.Cells((i + 1), 1) Then
      ' Determine close price
        closePrice = ws.Cells(i, 6).Value
      ' Yearly change calculation
        yearlyChange = closePrice - openPrice
      ' Percent
        If tempOpen > 0 Then
          percentChange = yearlyChange / openPrice
        Else
          percentChange = 0
        End If
      openPrice = ws.Cells(i + 1, 3).Value
      totalStockVol = totalStockVol + ws.Cells(i, 7).Value
      ' Inserting the values
        ws.Range("I" & summaryTableRow).Value = tickerSymbol
        ws.Range("J" & summaryTableRow).Value = yearlyChange
      ' Sending percentage value to row K
        ws.Range("K" & summaryTableRow).Value = percentChange
        ws.Range("K" & summaryTableRow).NumberFormat = "0.00%"
      ' Sending total stock volume to value l
        ws.Range("L" & summaryTableRow).Value = totalStockVol
       ' Color formatting
        If yearlyChange > 0 Then
        ws.Range("J" & summaryTableRow).Interior.ColorIndex = 4
        Else
        ws.Range("J" & summaryTableRow).Interior.ColorIndex = 3
        End If
        ' Keep stepping through table rows to fill out
        summaryTableRow = summaryTableRow + 1
        ' reseting values
        totalStockVol = 0
      Else
        totalStockVol = totalStockVol + Cells(i, 7).Value
      End If
    Next i
  Next ws
End Sub
