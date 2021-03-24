Attribute VB_Name = "Module1"
Sub createHeader(book)
    ' Create aggregate Table
    book.Range("I1").Value = "Ticker"
    book.Range("J1").Value = "Yearly Change"
    book.Range("K1").Value = "Percent Change"
    book.Range("L1").Value = "Total Stock Volume"
    
    ' Set column styles
    book.Columns("J:J").NumberFormat = ("0.00")
    book.Columns("K:K").NumberFormat = ("0.00%")
    
    ' Create max aggregate Table
    book.Range("P1").Value = "Ticker"
    book.Range("Q1").Value = "Value"
    book.Range("O2").Value = "Greatest % Increase"
    book.Range("O3").Value = "Greatest % Decrease"
    book.Range("O4").Value = "Greatest Total Volume"
    
    ' Set Cell styles
    book.Range("Q2:Q3").NumberFormat = ("0.00%")
    
End Sub


Sub recordStock(book, ticker, openPrice, closePrice, stockVolume, stockIndex)
    Dim priceChange As Double
    Dim percentChange As Double
    
    ' Calculate changes
    priceChange = closePrice - openPrice
    
    If openPrice > 0 Then
        percentChange = priceChange / openPrice
    Else
        percentChange = 0
    End If
    
    ' Record values in next available row
    book.Range("I" & stockIndex).Value = ticker
    book.Range("J" & stockIndex).Value = priceChange
    book.Range("K" & stockIndex).Value = percentChange
    book.Range("L" & stockIndex).Value = stockVolume
    
    ' Change priceChange column color
    If priceChange >= 0 Then
        book.Range("J" & stockIndex).Interior.ColorIndex = 4
    Else
        book.Range("J" & stockIndex).Interior.ColorIndex = 3
    End If
    
End Sub

Sub stockLooper(book)
    ' Variables for first and last row of sheet
    Dim firstRow As Long
    Dim lastRow As Long
    
    Dim ticker As String
    ' Open price of the stock at the beginning of the year
    Dim firstOpen As Double
    ' Close price of the stock at the end of the year
    Dim lastClose As Double
    ' Running total of the volume of stock traded throughout year
    Dim volumeTotal As Double
    
    firstRow = 2
    lastRow = book.Cells(book.Rows.Count, "A").End(xlUp).Row
    stockIndex = 1
    
    For Row = firstRow To lastRow
    
        ' Set new ticker values when a new ticker is reached
        If book.Cells(Row - 1, 1) <> book.Cells(Row, 1) Then
            ticker = book.Cells(Row, 1)
            firstOpen = book.Cells(Row, 3)
            volumeTotal = book.Cells(Row, 7)
            stockIndex = stockIndex + 1
        ' Add volume to total if ticker remains the same
        Else
            volumeTotal = volumeTotal + book.Cells(Row, 7)
        End If

        ' Set last value in ticker
        If book.Cells(Row + 1, 1) <> book.Cells(Row, 1) Then
            lastClose = book.Cells(Row, 6)
            'Record values to table
            Call recordStock(book, ticker, firstOpen, lastClose, volumeTotal, stockIndex)
        End If
    
    Next Row
        
  
End Sub

Sub aggregateLooper(book)
    ' Variables for greatest increase
    Dim giTicker As String
    Dim giValue As Double
    
    ' Variables for greatest decrease
    Dim gdTicker As String
    Dim gdValue As Double
    
    ' Variables for greatest volume
    Dim gvTicker As String
    Dim gvValue As Double
    
    Dim firstRow As Long
    Dim lastRow As Long
    
    giValue = 0
    gdValue = 0
    gvValue = 0
    
    firstRow = 2
    lastRow = book.Cells(book.Rows.Count, "I").End(xlUp).Row
    
    For Row = firstRow To lastRow
        
        If book.Range("K" & Row) > giValue Then
            giTicker = book.Range("I" & Row).Value
            giValue = book.Range("K" & Row).Value
        End If
        
        If book.Range("K" & Row) < gdValue Then
            gdTicker = book.Range("I" & Row).Value
            gdValue = book.Range("K" & Row).Value
        End If
        
        If book.Range("L" & Row) > gvValue Then
            gvTicker = book.Range("I" & Row).Value
            gvValue = book.Range("L" & Row).Value
        End If
        
    Next Row
    
    ' Set table values
    book.Range("P2").Value = giTicker
    book.Range("Q2").Value = giValue
    book.Range("P3").Value = gdTicker
    book.Range("Q3").Value = gdValue
    book.Range("P4").Value = gvTicker
    book.Range("Q4").Value = gvValue
        
    
End Sub

Sub stockChanges()
    
    For Each ws In Worksheets
    
    ' Create Header row for summation
    Call createHeader(ws)
    
    ' Loop through sheet data to aggregate data
    Call stockLooper(ws)
    
    ' Loop through aggregate data to find greatest values
    Call aggregateLooper(ws)
    
    Next ws
End Sub
