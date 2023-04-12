Sub stock_analysis():

Dim ws As Worksheet
    'Indicates more than 1 worksheet Loop
   For Each ws In Worksheets
   
   ws.Activate


'Declare variables
Dim i As Long
Dim Yearly_Change As Double
Dim j As Integer
Dim Ticker_Volume As Double
Dim Ticker_Count As Long
Dim LastRow As Long
Dim Percent_Change As Double
Dim days As Integer
Dim Greatest_Increase As Double
Dim Greatest_Decrease As Double
Dim Greatest_Total As Double

   

   'Add column and row headers
    ws.Range("I1,P1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
   
   'Initial values and Last Row
   j = 0
   Total_Volume = 0
   Yearly_Change = 0
   Ticker_Count = 2
   
   
   'Loop through all ticker names/rows
   LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
   For i = 2 To LastRow
   
   ' See if we are still within the same ticker sign
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
           'Stores results
           Total_Volume = Total_Volume + ws.Cells(i, 7).Value
           'Zero total volume
           If Total_Volume = 0 Then
               'put results in the fields
               ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
               ws.Range("J" & 2 + j).Value = 0
               ws.Range("K" & 2 + j).Value = "%" & 0
               ws.Range("L" & 2 + j).Value = 0
            Else
              'Find non zero starting value
              If ws.Cells(Ticker_Count, 3) = 0 Then
               For find_value = Ticker_Count To i
                       If ws.Cells(find_value, 3).Value <> 0 Then
                           Ticker_Count = find_value
                           Exit For
                       End If
                Next find_value
              End If
                     
            'Calculate change
            Yearly_Change = (Cells(i, 6) - Cells(Ticker_Count, 3))
            Percent_Change = Yearly_Change / Cells(Ticker_Count, 3)
            
            'Start of the next stock ticker and record results
            Ticker_Count = i + 1
            
            'Record results
            Range("I" & 2 + j).Value = Cells(i, 1).Value
            Range("J" & 2 + j).Value = Yearly_Change
            Range("J" & 2 + j).NumberFormat = "0.00"
            Range("K" & 2 + j).Value = Percent_Change
            Range("K" & 2 + j).NumberFormat = "0.00%"
            Range("L" & 2 + j).Value = Total_Volume
            
                'Apply conditional formatting to Yearly Change column, green if negative and red if positive
                If (Yearly_Change > 0) Then
                ws.Range("J" & 2 + j).Interior.ColorIndex = 4
                
                ElseIf (Yearly_Change <= 0) Then
                ws.Range("J" & 2 + j).Interior.ColorIndex = 3
                
                End If
               
               
            End If
                         
           'Move on to next ticker symbol and reset variables
           Total_Volume = 0
           Yearly_Change = 0
           j = j + 1
           days = 0
           'Add results for each ticker symbol
       Else
           Total_Volume = Total_Volume + ws.Cells(i, 7).Value
    End If
    
          
    Next i
    
    
    'Find the Max and Min of percent change and max volume
    ws.Range("Q2") = "%" & WorksheetFunction.Max(Range("K2:K" & LastRow)) * 100
    ws.Range("Q3") = "%" & WorksheetFunction.Min(Range("K2:K" & LastRow)) * 100
    ws.Range("Q4") = WorksheetFunction.Max(Range("L2:L" & LastRow))

   
    ' print ticker symbols for greatest increase and decrease and greatest volume
    ws.Range("P2") = ws.Cells(increase_number + 1, 9)
    ws.Range("P3") = ws.Cells(decrease_number + 1, 9)
    ws.Range("P4") = ws.Cells(volume_number + 1, 9)

    
   
    Next ws
    
        MsgBox ("Assignment completed for Rose Kasper")

End Sub