Public Sub Ticker()

    'Dim column As Long
    Dim lastRow As Double
   
    Dim Ticker As String
   
    'Track Ticker
    Dim Ticker_Tracker
   
    'Keep track of location
    Dim Ticker_Summary As Long
    Ticker_Summary = 1
   
    'Keep track of Yearly Change
    'Dim Stock As Double
    'Stock = 0
   
    'Keep track of opening day amount
    Dim Open_Day As Double
    Open_Day = 0
   
    'Keep track of Closing last day amount
    Dim Last_Day As Double
    Last_Day = 0
   
    'Store percentage
    Dim Percent As Double
   
   
    'Column = 1
   
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
   
    For i = 1 To lastRow
       
            'If Cells(1, 8).Value = Cells(1, 1).Value Then
           
            If Cells(i + 1, 1).Value = Cells(i, 1).Value Then
       
                'Store Opening Amount
                Last_Day = Range("F" & i + 1).Value
           
            End If
           
       
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
       
       
            'Add to Yearly Total
            'Stock = Stock + Cells(i + 1, 7).Value
       
            'Store Ticker
            Ticker = Cells(i, 1).Value
       
            'Store Ticker in far Column
            Range("H" & Ticker_Summary).Value = Ticker
           
            'Print Stock Summary
            'Range("K" & Ticker_Summary).Value = Stock
           
            'Store Last Day
            Open_Day = Cells(i + 1, 3)
           
            'Add to Ticker Summy
            Ticker_Summary = Ticker_Summary + 1
           
            'Reset Stock Total
            'Stock = 0
           
            'Store Last Day
            'Last_Day = Cells(i, 6)
           
            'Continue with loop
            Else
           
             If Open_Day = 0 Then
               
                Open_Day = 1
               
            End If
           
            If Last_Day = 0 Then
           
                Last_Day = 1
               
            End If
           
            'Move Yearly Change into Column I
            Range("I" & Ticker_Summary) = Last_Day - Open_Day
           
           
            'Move Percentage of Yearly Change into Column J
            Percent = ((Last_Day - Open_Day) / Open_Day)
           
            Range("J" & Ticker_Summary) = FormatPercent(Percent)
           
            'If Percent > 0 Then
       
                'Cells(Ticker_Summary, 10).Interior.ColorIndex = 4
           
            'Else
           
                'Cells(Ticker_Summary, 10).Interior.ColorIndex = 3
   
            'End If
           
            'Add to the Stock Total
            'Stock = Stock + Cells(i + 1, 7).Value
           
   
        End If
       
    Next i

End Sub

Public Sub Format_Headers()

Cells(1, 8) = "Ticker"

Cells(1, 9) = "Yearly Change"

Cells(1, 10) = "Percent Change"

Cells(1, 11) = "Total Stock Volume"

Cells(4, 14) = "Greatest % Increase"

Cells(5, 14) = "Greatest % Decrease"

Cells(6, 14) = "Greatest Total Volume"

Cells(3, 15) = "Ticker"

Cells(3, 16) = "Value"


End Sub



Public Sub TotalVolume()

'Store a value for Stock Volume
Dim rng As Range
Dim H_S As Double

'Get last row
lastRow = Cells(Rows.Count, 1).End(xlUp).Row


'set range to get largest value
Set rng = ActiveSheet.Range("K1:K100000")

'Find largest stock total
H_S = Application.WorksheetFunction.Max(rng)

Cells(6, 16) = H_S

For i = 1 To lastRow

    If Cells(i + 1, 11).Value = H_S Then
   
    'Return Ticker
   
    Cells(6, 15) = Cells(i + 1, 8)
   
    End If
   
    Next i
   
End Sub


Public Sub PercentIncrease()

'Store a value for Stock Volume
Dim rng As Range
Dim P_I As Double

'Get last row
lastRow = Cells(Rows.Count, 1).End(xlUp).Row


'set range to get largest value
Set rng = ActiveSheet.Range("J1:J100000")

'Find largest stock total
P_I = Application.WorksheetFunction.Max(rng)

Cells(4, 16) = FormatPercent(P_I)

For i = 1 To lastRow

    If Cells(i + 1, 10).Value = P_I Then
   
    'Return Ticker
   
    Cells(4, 15) = Cells(i + 1, 8)
   
    End If
   
    Next i
   
End Sub

Public Sub PercentDecrease()

'Store a value for Stock Volume
Dim rng As Range
Dim P_D As Double

'Get last row
lastRow = Cells(Rows.Count, 1).End(xlUp).Row


'set range to get largest value
Set rng = ActiveSheet.Range("J1:J30000")

'Find largest stock total
P_D = Application.WorksheetFunction.Min(rng)

Cells(5, 16) = FormatPercent(P_D)

For i = 1 To lastRow

    If Cells(i + 1, 10).Value = P_D Then
   
    'Return Ticker
   
    Cells(5, 15) = Cells(i + 1, 8)
   
    End If
   
    Next i
   
End Sub


Public Sub FormatAllSheets()

Dim i As Integer

        i = 1

        Do While i <= Worksheets.Count
            Worksheets(i).Select
             
            Ticker
            Opening_Day_Difference
            Stock_Total
            Format_Headers
            TotalVolume
            PercentIncrease
            PercentDecrease
           
            i = i + 1
        Loop


End Sub

Public Sub Opening_Day_Difference()


lastRow = Cells(Rows.Count, 10).End(xlUp).Row

For i = 1 To lastRow
   
            
    If Cells(i + 1, 10).Value > 0 Then
           
       
                Cells(i + 1, 10).Interior.ColorIndex = 4
           
            Else
           
                Cells(i + 1, 10).Interior.ColorIndex = 3
   
            End If

       
    Next i


End Sub

Public Sub StockTotal()

'Keep track of location
    Dim Ticker_Summary As Long
    Ticker_Summary = 1

'Keep track of Yearly Change
    Dim Stock As Double
    Stock = 0
    
    
lastRow = Cells(Rows.Count, 1).End(xlUp).Row
   
    For i = 1 To lastRow
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
       
            'Add to Yearly Total
            Stock = Stock + Cells(i + 1, 7).Value
            
            'Print Stock Summary
            Range("K" & Ticker_Summary).Value = Stock
            
            'Add to Ticker Summy
            Ticker_Summary = Ticker_Summary + 1
            
            'Reset Stock Total
            Stock = 0
        
         Else
            
            'Add to the Stock Total
            Stock = Stock + Cells(i + 1, 7).Value
    
    End If
    
   Next i

End Sub
