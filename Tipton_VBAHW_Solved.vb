Sub Stonks():
 
 'Run through all worksheets
 For Each WS In ActiveWorkbook.Worksheets
 WS.Activate
 
 
 Dim sht As Worksheet
    Set sht = ActiveSheet
    LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
 Dim TickerName As String
 Dim YearlyChange As Double
 Dim PercentChange As Double
 Dim StockVolume As Double
    StockVolume = 0
 Dim OpenPrice As Double
 Dim ClosePrice As Double
 Dim Row As Double
    Row = 2
 
  
  'Create Headers
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Range("O2").Value = "Greatest % Change"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    
  'Set Open Price
  OpenPrice = Cells(2, 3).Value
  
  
  For i = 2 To LastRow
  
       
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        
        'Set Ticker Name
        TickerName = Cells(i, 1).Value
        Cells(Row, 9).Value = TickerName
        
        'Set Close Price
        ClosePrice = Cells(i, 6).Value
        
        'Find Yearly Change
        YearlyChange = ClosePrice - OpenPrice
        Cells(Row, 10).Value = YearlyChange
        Cells(Row, 10).NumberFormat = "0.000000000"
        
        'Find Percent Change
            If (OpenPrice = 0 And ClosePrice = 0) Then
                PercentChange = 0
            ElseIf (OpenPrice = 0 And ClosePrice <> 0) Then
                PercentChange = 1
            Else
                PercentChange = YearlyChange / OpenPrice
                Cells(Row, 11).Value = PercentChange
                Cells(Row, 11).NumberFormat = "0.00%"
            End If
            
        'Find Total Volume
        StockVolume = StockVolume + Cells(i, 7).Value
        Cells(Row, 12).Value = StockVolume
        
        Row = Row + 1
        
        'Reset Open Price
        OpenPrice = Cells(i + 1, 3).Value
        
        'Reset Total Stock Volume
        StockVolume = 0
        
        Else
        StockVolume = StockVolume + Cells(i, 7).Value
        
        End If
    
    
    Next i
    
    'Establish Last Row in Yearly Change
    YCLastRow = WS.Cells(Rows.Count, 10).End(xlUp).Row
    
    'Conditional Formatting
    For j = 2 To YCLastRow
        If Cells(j, 10).Value >= 0 Then
        Cells(j, 10).Interior.ColorIndex = 4
        Else
        Cells(j, 10).Interior.ColorIndex = 3
        End If
    Next j
    
    'Establish Last Row in Percent Change
    PCLastRow = WS.Cells(Rows.Count, 11).End(xlUp).Row
    
    'Find Greatest % Increase Value
    For h = 2 To PCLastRow
    If Cells(h, 11).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & PCLastRow)) Then
    Cells(2, 17).Value = Cells(h, 11).Value
    Cells(2, 17).NumberFormat = "0.00%"
    'Find Greatest % Increase Ticker
    Cells(2, 16).Value = Cells(h, 9).Value
    
  
    'Find Greatest % Decrease Value
    ElseIf Cells(h, 11).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & PCLastRow)) Then
    Cells(3, 17).Value = Cells(h, 11).Value
    Cells(3, 17).NumberFormat = "0.00%"
    'Find Greatest % Decrease Ticker
    Cells(3, 16).Value = Cells(h, 9).Value
    
    'Find Greatest Total Volume
    ElseIf Cells(h, 12).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & PCLastRow)) Then
    Cells(4, 17).Value = Cells(h, 12).Value
    'Find Greatest Total Volume Ticker
    Cells(4, 16).Value = Cells(h, 9).Value
    
    
    End If
    
    Next h
    
Next WS

End Sub


