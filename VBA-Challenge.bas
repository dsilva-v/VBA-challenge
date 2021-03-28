Attribute VB_Name = "Module1"
Sub Ticker():

Dim ws As Worksheet
Dim LastRow As Long
Dim TotalStockVolume As Double
Dim summaryRowTable As Long
Dim YearlyOpen As Double
Dim YearlyClose As Double
Dim YearlyChange As Double
Dim PercentageChange As Double
Dim tickerChangeIndex As Long


'Loop through each worksheet
For Each ws In Worksheets
      'Set all the inital values needed for a worksheet before looping starts
      'Initialize the LastRow with the last row with data in the column 1
      'Initalize the TotalStockVolume with 0 (used as a count variable)
      'Initalize the YearlyOpen with 0
      'Initalize the YearlyClose with 0
      'Initalize the summaryTableRow with 2(to start looping up the ticker from row 2 in excel)
      'Initialize the headers for I, J, K, L columns
      'Initialize the ticker Index to keep track of YearlyClose, YearlyOpen for each ticker
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        TotalStockVolume = 0
         YearlyOpen = 0
         YearlyClose = 0
         summaryTableRow = 2
         ws.Range("I1").Value = "Ticker"
         ws.Range("J1").Value = "Yearly Change"
         ws.Range("K1").Value = "Percent Change"
         ws.Range("L1").Value = "Total Stock Volume"
         tickerChangeIndex = 2
            
'Loop through each row  with data in column1 in the individual worksheet
For i = 2 To LastRow
       'If the  tickers in the adjacent cells are not same then
       'Set the ticker Name
       'Calcualte YearlyChange while preserving YearlyOpen
       If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
          
        ws.Range("I" & summaryTableRow).Value = ws.Cells(i, 1).Value
    
       YearlyOpen = ws.Cells(tickerChangeIndex, 3).Value
        YearlyClose = ws.Cells(i, 6).Value
        YearlyChange = YearlyClose - YearlyOpen
       
        ws.Range("J" & summaryTableRow).Value = YearlyChange
        
        'conditional format of the yearly change with colors
         If YearlyChange >= 0 Then
          
         ws.Range("J" & summaryTableRow).Interior.ColorIndex = 4
         Else
         ws.Range("J" & summaryTableRow).Interior.ColorIndex = 3
         End If
       
        'Handling error scenario where YearlyOpen is 0
        If YearlyOpen = 0 Then
        PercentChange = YearlyChange
        
        
        Else
        PercentChange = YearlyChange / YearlyOpen
        
        End If
        
        ws.Range("K" & summaryTableRow).Value = PercentChange
          'Format Percent change to show % symbol
        ws.Range("K" & summaryTableRow).NumberFormat = "0.00%"
       
        ws.Range("L" & summaryTableRow).Value = TotalStockVolume + ws.Cells(i, 7).Value
        summaryTableRow = summaryTableRow + 1
        TotalStockVolume = 0
         tickerChangeIndex = i + 1

         Else
         TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
          End If
          
          
          Next i
          

          Next ws
End Sub

