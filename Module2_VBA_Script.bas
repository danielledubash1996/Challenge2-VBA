Attribute VB_Name = "Module1"
Sub Module2()

For Each ws In Worksheets
    'Name of Title Rows
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"
    
    'setting varibale dimensions
    Dim ticker As String
    Dim yearlychange As Double
    Dim percentchange As Double
    Dim StockVolume As Double
    
    Dim firstopen As Double
    Dim lastclose As Double
    
    Dim lastrow As Double
    
    Dim num As Double
    Dim total As Double
    
    'Setting Variables
    StockVolume = 0
    num = 2
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    firstopen = ws.Range("C2")
    
    For i = 2 To lastrow
        'Formulas
        StockVolume = StockVolume + ws.Cells(i, 7).Value
        ticker = ws.Cells(i, 1)
        lastopen = ws.Cells(i, 6)
        yearlychange = lastopen - firstopen
        percentchange = yearlychange / firstopen
    
    'If statement to print out the values of them columns
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
            ws.Cells(num, 9) = ticker
            ws.Cells(num, 10) = yearlychange
            ws.Cells(num, 11) = FormatPercent(percentchange)
            ws.Cells(num, 12) = StockVolume
            
        'reset variable for next loop
        firstopen = ws.Cells(i + 1, 3)
        num = num + 1
        StockVolume = 0
        End If
        

    Next i
    
    'Conditional Formatting
 
    For j = 2 To lastrow
    Dim ranges As Range
    Set ranges = ws.Range("J" & j)
        
        If ranges > 0 Then
        ranges.Interior.ColorIndex = 4
        
        ElseIf ranges < 0 Then
        ranges.Interior.ColorIndex = 3
        
        End If
        
        Next j
        
   'Adding Functionality
   
   'Name of Title Rows
    ws.Range("P1") = "Ticker"
    ws.Range("Q1") = "Value"
    ws.Range("O2") = "Greatest % Increase"
    ws.Range("O3") = "Greatest % Decrease"
    ws.Range("O4") = "Greatest Total Volume"
    
    'setting variable dimensions
    Dim increase As Double
    Dim decrease As Double
    Dim volume As Double
    Dim m As Double
    
    'Defining Ranges
    change = ws.Range("K:K")
    totalvolume = ws.Range("L:L")
    
    'Setting Variables
    increase = WorksheetFunction.Max(change)
    decrease = WorksheetFunction.Min(change)
    volume = WorksheetFunction.Max(totalvolume)
    
    'Setting For Loop to Look through and grab the right valies
    For m = 2 To lastrow
    
        If ws.Cells(m, 11).Value = increase Then
            ws.Range("P2") = ws.Cells(m, 9)
            ws.Range("Q2") = ws.Cells(m, 11)
            
        ElseIf ws.Cells(m, 11).Value = decrease Then
            ws.Range("P3") = ws.Cells(m, 9)
            ws.Range("Q3") = ws.Cells(m, 11)
            
        ElseIf ws.Cells(m, 12).Value = volume Then
            ws.Range("P4") = ws.Cells(m, 9)
            ws.Range("Q4") = ws.Cells(m, 12)
        End If
        
    Next m
        
   Next ws

End Sub
