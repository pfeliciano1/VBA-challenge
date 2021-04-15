Sub Stocks()

'Variable to cycle the Worksheets
Dim ws As Worksheet
'Variable for ticker symbol
Dim ticker As String
'Variable for total volume stock
Dim total As Long
'Variable for the Summary table
Dim SummaryTable As Integer
'Variable for the Loop through tickers
Dim i As Long
'Variable for the Loop for greatest values
Dim j As Long
'Variable for open year value
Dim OpenYr As Double
'Variable for close year value
Dim CloseYr As Double
'Variable for year change value
Dim YrChange As Double
'Variable for percent change value
Dim PercentChange As Double

'Loop through Worksheets
For Each ws In Worksheets

'Summary Table Headers
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

'Start values for each variable
total = 0
SummaryTable = 2
OpenYr = 0
CloseYr = 0
YrChange = 0
PercentChange = 0

'Declare variable for the loop to cycle through all rows
Dim LastRow As Long
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Loop through Tickers
For i = 2 To LastRow

'    'Conditional for the ticker symbol and the values
    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        'Set Ticker name
        ticker = ws.Cells(i, 1).Value
        'Add to total stock volume
        total = total + ws.Cells(i, 7).Value
        'Print ticker in Summary Table
        ws.Range("I" & SummaryTable).Value = ticker
        'Print total stock volume in Summary Table
        ws.Range("L" & SummaryTable).Value = total
        
 '       OpenYr = ws.Cells(i, 3).Value
        CloseYr = ws.Cells(i, 6).Value
        'Year change formula and value in Summary Table
        YrChange = CloseYr - OpenYr
        ws.Range("J" & SummaryTable).Value = YrChange
            
  '          'Conditional for year change and the formatting
            If YrChange > 0 Then
                ws.Range("J" & SummaryTable).Interior.ColorIndex = 4
            
  '          ElseIf YrChange = 0 Then
                ws.Range("J" & SummaryTable).Interior.ColorIndex = 0
            
   '         ElseIf YrChange < 0 Then
                ws.Range("J" & SummaryTable).Interior.ColorIndex = 3
            
  '          End If
            
  '          'Conditional to Calculate Percent Change
            If OpenYr <> 0 Then
            PercentChange = YrChange / OpenYr  'Percent Change formula
            
 '           Else
            PercentChange = 0
            
  '          End If
            
  '          ws.Range("K" & SummaryTable).Value = PercentChange
            ws.Range("K" & SummaryTable).NumberFormat = "0.00%"
            
  '      SummaryTable = SummaryTable + 1
        total = 0
    End If
    
        
Next i

'BONUS
'Greatest values Table Headers
ws.Range("O2").Value = "Greatest % Increased"
ws.Range("O3").Value = "Greatest % Decreased"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"


'Determine the Last Row
LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row

'Loop through Tickers to identify Greatest values
For j = 2 To LastRow

'Conditional to identify Greatest % increase
    If ws.Cells(j, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & LastRow)) Then
        ws.Range("P2").Value = ws.Cells(j, 9).Value
        ws.Range("Q2").Value = ws.Cells(j, 11).Value
        ws.Range("Q2").NumberFormat = "0.00%"
        
'Conditional to identify Greatest % decrease
    ElseIf ws.Cells(j, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & LastRow)) Then
        ws.Range("P3").Value = ws.Cells(j, 9).Value
        ws.Range("Q3").Value = ws.Cells(j, 11).Value
        ws.Range("Q3").NumberFormat = "0.00%"
        
 'Conditional to identify Greatest Total Volume
    ElseIf ws.Cells(j, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & LastRow)) Then
        ws.Range("P4").Value = ws.Cells(j, 9).Value
        ws.Range("Q4").Value = ws.Cells(j, 12).Value
        
 '   End If
    
Next j
'Format the columns with AutoFit
ws.Columns("I:L").AutoFit
ws.Columns("O:Q").AutoFit

Next ws

End Sub
