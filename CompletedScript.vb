

Sub AllSheets()
' Script to run the stock stats subroutine

    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call Stock_stats
    Next
    Application.ScreenUpdating = True
End Sub

Sub Stock_stats()

'Set the initial variable for ticker
Dim ticker As String

'Set an initial variable for holding the total per credit card brand

'Dim Yearly_Change As Double
'Dim Percentage_Change As Double
'Dim MaxPercentChange As Double
'Dim MinPercentChange As Double
'Dim HighVolume As Long


Dim RowCount As Long
Dim SumCount As Long
Dim FirstRow As Long
Dim LastRow As Long


' Keep track of the location for each ticker in the summary table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

Dim Ticker_Total As Long

'Find the last non-blank cell in column A(1)
RowCount = Cells(Rows.Count, 1).End(xlUp).Row
    
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percentage"
Cells(1, 12).Value = "Total Stock Volume"

    
' Loop through all data
For i = 2 To RowCount

' Check to see if this was the same ticker reference
  If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        ' Set the Ticker name
        ticker = Cells(i, 1).Value

        'Set First Row and Reset the Ticket Counter
        FirstRow = i - Ticker_Total
        LastRow = FirstRow + Ticker_Total
        ' Add to the Ticker total volume
        'StockVolume = WorksheetFunction.Sum(Range(Cells(i, 7), Cells(LastRow, 7)))

        ' Print the ticker name in the Summary table
        Range("I" & Summary_Table_Row).Value = ticker
        
        ' Print the Summary Table
                Range("J" & Summary_Table_Row).Formula = "=(F" & LastRow & "- C" & FirstRow & ")"
                Range("L" & Summary_Table_Row).Formula = "=Sum(G" & FirstRow & ":G" & LastRow & ")"
                
        ' To assign a value for any ticker with 0 values
                If Range("C" & FirstRow).Value = 0 Or Range("F" & LastRow).Value = 0 Then
                    Range("K" & Summary_Table_Row).Value = 0
                Else
                    Range("K" & Summary_Table_Row).Formula = "=((F" & LastRow & "- C" & FirstRow & ")/C" & FirstRow & ")"
                End If
                
        ' Change the Nunber format to display Percentage
        
                Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                
        ' To Change cell colour based on the value
        
                If Range("J" & Summary_Table_Row).Value < 0# Then
                    Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                Else
                    Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                End If
                
        ' Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
        LastRow = 0
        Ticker_Total = 0
     ' If the cell immediately following a row is the same ticker
    Else
        ' Count the number of rows for each of the Stock
        Ticker_Total = Ticker_Total + 1
    End If
  Next i
  
  ' Create the additional summary table headings
  
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
  
  ' Count the number of rows of the summary table
  
    SumCount = Cells(Rows.Count, 11).End(xlUp).Row
                
  ' Find the Maximum Percentage Change
    Range("Q2").Formula = "=Max(K2:K" & SumCount & ")"
  ' Find the minimum Percentage Change
    Range("Q3").Formula = "=Min(K2:K" & SumCount & ")"
  ' Find the Maximum Total Volume
    Range("Q4").Formula = "=Max(L2:L" & SumCount & ")"
    
  ' Find the ticker reference for each of the above
    Range("P2").Formula = "=INDEX(I2:I" & SumCount & ",MATCH(Q2,K2:K" & SumCount & ",0))"
    Range("P3").Formula = "=INDEX(I2:I" & SumCount & ",MATCH(Q3,K2:K" & SumCount & ",0))"
    Range("P4").Formula = "=INDEX(I2:I" & SumCount & ",MATCH(Q4,L2:L" & SumCount & ",0))"
    
End Sub

