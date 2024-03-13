Attribute VB_Name = "Module11"
Sub stocks()

Dim ws As Integer
Dim rows As Long
Dim LastRow As Long
Dim StockCount As Integer
Dim TotalStockVol As Variant
Dim TotalStockVolMax As Variant
Dim TotalVolMaxTckr As String
Dim GreatPercIncTckr As String
Dim GreatPercDecTckr As String
Dim OpeningDate As Long ' change to Date if original data is in date format
Dim OpeningDateRow As Long
Dim ClosingDate As Long
Dim ClosingDateRow As Long
Dim OP As Single
Dim CP As Single
Dim YrPercChange As Single
Dim GreatPercInc As Single
Dim GreatPercDec As Single

Worksheets(1).Activate

Debug.Print OpeningDate, TotalStockVol
    
For ws = 1 To 3     'tabs 2018, 2019, 2020

    ThisWorkbook.Worksheets(ws).Activate
    rows = 2
    LastRow = Range("A1").End(xlDown).Row
    Debug.Print LastRow
    StockCount = 1
    OpeningDate = Cells(rows, 2)
    OpeningDateRow = 2
    TotalStockVol = Cells(rows, 7)
    GreatPercInc = 0
    GreatPercDec = 0
    TotalStockVolMax = 0
    GreatPercIncTckr = 0
    GreatPercDecTckr = 0
    TotalVolMaxTckr = 0
    YrPercChange = 0
    Worksheets(ws).Cells(1, 9) = "Ticker"
    Worksheets(ws).Cells(1, 10) = "Yearly Change"
    Worksheets(ws).Cells(1, 11) = "Percent Change"
    Worksheets(ws).Cells(1, 12) = "Total Stock Volume"
    
    For rows = 2 To LastRow
    'For rows = 2 To 2000
        If Cells(rows, 1) = Cells(rows + 1, 1) Then
        
            TotalStockVol = TotalStockVol + Cells(rows + 1, 7)
            If Cells(rows, 2) < OpeningDate Then
                OpeningDate = Cells(rows, 2)
                OpeningDateRow = rows
            End If
            If Cells(rows + 1, 2) > Cells(rows, 2) Then
                ClosingDate = Cells(rows + 1, 2)
                ClosingDateRow = rows + 1
            End If
            'Debug.Print OpeningDate, OpeningDateRow
        Else
        
            StockCount = StockCount + 1
            OP = Cells(OpeningDateRow, 3).Value 'opening price of the year
            CP = Cells(ClosingDateRow, 6).Value 'closing price of the year
            YrPercChange = (CP - OP) / OP
            Worksheets(ws).Cells(StockCount, 9) = Cells(rows, 1)
            Worksheets(ws).Cells(StockCount, 10) = Format(CP - OP, "#.00")
            Worksheets(ws).Cells(StockCount, 11) = Format(YrPercChange, "#.00%")
            Worksheets(ws).Cells(StockCount, 12) = TotalStockVol
        
            If Worksheets(ws).Cells(StockCount, 10) > 0 Then
            
                Range(Worksheets(ws).Cells(StockCount, 10), Worksheets(ws).Cells(StockCount, 11)).Interior.Color = RGB(0, 255, 0)
                
            ElseIf Worksheets(ws).Cells(StockCount, 10) < 0 Then
            
                Range(Worksheets(ws).Cells(StockCount, 10), Worksheets(ws).Cells(StockCount, 11)).Interior.Color = RGB(255, 0, 0)
            
            End If
            
            TotalStockVol = Cells(rows + 1, 7)
            OpeningDate = Cells(rows + 1, 2)
            OpeningDateRow = rows + 1
            ClosingDateRow = rows
            
        End If
        ' Master summary analysis:
        If YrPercChange > GreatPercInc Then
                GreatPercInc = YrPercChange
                GreatPercIncTckr = Cells(rows, 1)
        End If
        If YrPercChange < GreatPercDec Then
                GreatPercDec = YrPercChange
                GreatPercDecTckr = Cells(rows, 1)
        End If
        If TotalStockVol > TotalStockVolMax Then
                TotalStockVolMax = TotalStockVol
                TotalVolMaxTckr = Cells(rows, 1)
        End If
    
    Next rows

    Cells(1, 16) = "Ticker"
    Cells(1, 17) = "Value"
    Cells(2, 15) = "Greatest % Inrease"
    Cells(2, 16) = GreatPercIncTckr
    Cells(2, 17) = Format(GreatPercInc, "#.00%")
    Cells(3, 15) = "Greatest % Decrease"
    Cells(3, 16) = GreatPercDecTckr
    Cells(3, 17) = Format(GreatPercDec, "#.00%")
    Cells(4, 15) = "Greatest Total Volume"
    Cells(4, 16) = TotalVolMaxTckr
    Cells(4, 17) = TotalStockVolMax
    Worksheets(ws).Columns("J:O").AutoFit
Next ws

Worksheets(1).Activate

End Sub
