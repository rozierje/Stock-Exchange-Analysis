Attribute VB_Name = "Module5"
Sub stock_exchange()

        'declarations
        Dim ticker As Integer
                ticker = 0
        Cells(1, 9).Value = "Ticker Symbol"
        Cells(1, 9).ColumnWidth = 15
        'find the last row of the data set
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        'MsgBox (lastrow)
        'place first ticker symbol
        Range("I2").Value = Range("A2").Value
    
        'find other ticker symbols
        For i = 2 To lastrow
    
    
        'conditions
            If (Cells(i, 1).Value <> Cells(i + 1, 1)) Then
                ticker = (ticker + 1)
                Cells(2 + ticker, 9).Value = Cells(i + 1, 1)
            
                End If
            If (Cells(i, 1).Value = Cells(i + 1, 1)) Then
            'increase ticker count
            ticker = ticker
                End If
        Next i
    'MsgBox "Stock Symbols Placed."
    
'------------------------------------------------

    'declarations
        Dim min As Double
        Dim max As Double
        Dim day As Long
                day = 0
        'reset ticker
            ticker = 2
       'create headings
     Cells(1, 10).Value = "Yearly Price +/-"
        Cells(1, 10).ColumnWidth = 15
     Cells(1, 11).Value = "Yearly Price % Change"
        Cells(1, 11).ColumnWidth = 20
        'find last row of data set
      lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    'set parameters
    For i = 2 To lastrow
        'conditionals
    If Cells(i, 1).Value = Cells(i + 1, 1).Value Then
   
            day = (day + 1)
    Else
        If max <> 0 Then
    'declare minimum value
        min = Cells(i - day, 3).Value
   'declare maximum value
        max = Cells(i, 6).Value
    'add value to yearly change column
        Cells(ticker, 10).Value = (max - min)
    'add value to percent column
      
            Cells(ticker, 11).Value = (min / max / 100)
            Cells(ticker, 11).NumberFormat = "0.00%"
       End If
       If max = 0 Then
       End If
    
      'update ticker
        ticker = ticker + 1
        End If
    Next i
    'MsgBox "Yearly Price Change Entered."
    'MsgBox "Yearly Percentage Calculated."

'---------------------------------------------------------------
    lastrow = Cells(Rows.Count, 10).End(xlUp).Row
    
    For i = 2 To lastrow
    
    If Cells(i, 10).Value > 0 Then
        Cells(i, 10).Interior.ColorIndex = 43
    Else
        Cells(i, 10).Interior.ColorIndex = 1
        End If
    If Cells(i, 10).Value < 0 Then
        Cells(i, 10).Interior.ColorIndex = 3
        End If
        
    Next i
'--------------------------------------------------------------------------
'declarations
        Dim stock_volume As Double
        'reset day
                day = 0
        'reset ticker
            ticker = 2
       'create headings
     Cells(1, 12).Value = "Stock Volume Totals"
        Cells(1, 12).ColumnWidth = 18
        'find last row of data set
      lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    'set parameters
    stock_volume = 0
    
    For i = 2 To lastrow
        
        'when the data set does not match the current ticker symbol
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
   'build  and place stock volume
           stock_volume = stock_volume + Cells(i, 7).Value
            Cells(ticker, 12).Value = stock_volume
     'reset stock_volume
           stock_volume = 0
           ticker = ticker + 1
    Else
    'build stock volume
        stock_volume = stock_volume + Cells(i, 7).Value

        End If
    Next i
   ' MsgBox "Stock Volume Totals Entered."

'--------------------------------------------------------------------------
 
    'MsgBox "Stock Exchange Information Complete."
'----------------------------------------------------------
    'declarations
    Dim percent_low As Double
        percent_low = 0
    Dim percent_high As Double
        percent_high = 0
    'data set headings
    Range("n2").Value = "Greatest % Increase"
    Cells(1, 14).ColumnWidth = 20
    Cells(2, 16).NumberFormat = "0%"
    Cells(3, 16).NumberFormat = "0%"
    Range("n3").Value = "Greatest % Decrease"
    Range("n4").Value = "Greatest Volume"
    'find last row
     lastrow = Cells(Rows.Count, 10).End(xlUp).Row
'set comparison range
For i = 2 To lastrow
        'conditions based on zero
    If Cells(i, 10).Value <> 0 Then
       'finding largest drop in percentage
        If Cells(i, 10).Value <> Cells(i + 1, 10).Value Then
        'compare cell value to percent_low
            If Cells(i, 10).Value <> percent_low Then
                If Cells(i, 10).Value > percent_low Then
            'keep percent_low if lower than current value
                    percent_low = percent_low
            'update and display percent_low if higher than current value
                Else
                    percent_low = Cells(i, 10).Value
                    Range("p3").Value = percent_low
                    Range("o3").Value = Cells(i, 9).Value
                    
                    
                End If
            End If
        'compare cells value to percent_high
            If Cells(i, 10).Value <> percent_high Then
                If percent_high > Cells(i, 10).Value Then
                'keep percent_high if higher than current value
                    percent_high = percent_high
                'update and display percent_high if lower than current value
                Else
                    percent_high = Cells(i, 10).Value
                    Range("p2").Value = percent_high
                    Range("o2").Value = Cells(i, 9).Value

                End If
            End If
            
        End If
    End If
    Next i
'----------------------------------------------------------------
'declarations
    Dim most_volume As Double
        most_volume = 0
        Cells(1, 16).ColumnWidth = 10
 'find last row of data set for comparison
    lastrow = Cells(Rows.Count, 12).End(xlUp).Row
'set parameters loop
For i = 2 To lastrow
'compare current value to zero; or existing most_volume

    If Cells(i, 12).Value <> most_volume Then
'only looking for numbers larger than most_volume
        If Cells(i, 12).Value > most_volume Then
            most_volume = Cells(i, 12).Value
'display and update most_volume
            Range("p4").Value = most_volume
            Range("o4").Value = Cells(i, 9).Value
        End If
    End If
    
    Next i
    
        MsgBox "Stock Exchange Information Complete."
End Sub

