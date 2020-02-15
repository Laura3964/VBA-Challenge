Sub alphabet_testing()
 
 'Declare variables and set initial
    
    Dim Percent_Change As Double
    Dim Start As Long
    Dim Yearly_Change As Double
    Dim Ticker_Name As String
    Dim Ticker_Total_Volume As Double
    Dim Summary_Table_Row As Long
    Dim LastRow As Long
    Dim new_start As Long
    
 '_________________________________________________________________________
 
 For Each ws In Worksheets
 
 
    'The first row of a new ticker
    Start = 2
        
    Yearly_Change = 0
    
      
    ' Set an initial variable for holding the total stock volume per ticker
    
    Ticker_Total_Volume = 0

    ' Keep track of the location for each ticker in the summary table
    
    Summary_Table_Row = 2
    
    'Finding the last row of each worksheet
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    MsgBox ("The Last Row is " + Str(LastRow))
    
    
    
'____________________________________________________________________________

'Set Headers for summary table
ws.Range("K1") = "Ticker"
ws.Range("L1") = "Yearly Change"
ws.Range("M1") = "Percent Change"
ws.Range("N1") = "Total Stock Volume"


  ' Loop through all the tickers
  For i = 2 To LastRow

    ' Check if we are still within the same ticker name, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the ticker name
      Ticker_Name = ws.Cells(i, 1).Value

      ' Add to the Ticker_Total_Volume
      Ticker_Total_Volume = Ticker_Total_Volume + ws.Cells(i, 7).Value

      ' Print the Ticker Name in the Summary Table
      ws.Range("K" & Summary_Table_Row).Value = Ticker_Name

      ' Print the Brand Amount to the Summary Table
      ws.Range("N" & Summary_Table_Row).Value = Ticker_Total_Volume
       
      'To find first non-zero start value for opening price
      
      If ws.Cells(Start, 3).Value = 0 Then
        For new_start = Start To LastRow
            If ws.Cells(new_start, 3) <> 0 Then
                Start = new_start
                Exit For
            End If
                 
        Next new_start
      
      End If
      
                     
      'Yearly Change: last price of the year - first price of the year
      Yearly_Change = ws.Cells(i, 6).Value - ws.Cells(Start, 3).Value
      ws.Range("L" & Summary_Table_Row).Value = Yearly_Change
      
      'Percent Change: (Yearly Change/Start value)*100
      Percent_Change = Round((Yearly_Change / (ws.Cells(Start, 3).Value)) * 100, 2)
      ws.Range("M" & Summary_Table_Row).Value = Percent_Change
      
      'conditional format apply
      If Yearly_Change > 0 Then
        'green
        ws.Range("L" & Summary_Table_Row).Interior.ColorIndex = 4
                
      ElseIf Yearly_Change < 0 Then
        'red
        ws.Range("L" & Summary_Table_Row).Interior.ColorIndex = 3
        
      Else
        'no color
        ws.Range("L" & Summary_Table_Row).Interior.ColorIndex = 0
      
      End If
      
    
    '______________________________________________________________________
    
    'Next summary row information
    
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      Start = i + 1
      
      ' Reset the Ticker Total Volume
       Ticker_Total_Volume = 0
    '_____________________________________________________________________
    
    
    ' If the cell immediately following a row is the same ticker...
    Else

      ' Add to the Ticker Total Volume
      Ticker_Total_Volume = Ticker_Total_Volume + ws.Cells(i, 7).Value

    End If

  Next i
'_______________________________________________________________________

'Formatting
ws.Columns.AutoFit

Next ws


End Sub


