Sub stockFilter()

For Each ws In Worksheets

    Dim openedAt As Double
    Dim closedAt As Double
    Dim totalVolume As Double
    Dim yearlyChange As Double
    Dim nextRow As Integer
    Dim percentChange As Double
    Dim LastRow As Double
    Dim greatestTotalVolume As Double
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim gtvTicker As String
    Dim giTicker As String
    Dim gdTicker As String

    greatestTotalVolume = 0
    greatestIncrease = 0
    greatestDecrease = 0
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    openedAt = ws.Cells(2, 3)
    yearlyChange = 0
    nextRow = 2
    totalVolume = 0
    ws.Cells(1, 9) = "Ticker"
    ws.Cells(1, 10) = "Total Stock Volume"
    ws.Cells(1, 11) = "Yearly Change"
    ws.Cells(1, 12) = "Percent Change"

    For n = 2 To LastRow
        If ws.Cells(n + 1, 1) <> ws.Cells(n, 1) Then
            closedAt = ws.Cells(n, 6)
            yearlyChange = closedAt - openedAt
            ws.Cells(nextRow, 11) = yearlyChange
                If yearlyChange > 0 Then
                    ws.Cells(nextRow, 11).Interior.ColorIndex = 4
                ElseIf yearlyChange < 0 Then
                    ws.Cells(nextRow, 11).Interior.ColorIndex = 3
                End If
            percentChange = yearlyChange / openedAt
            ws.Cells(nextRow, 12) = percentChange
            ws.Cells(nextRow, 12).NumberFormat = "0.00%"
                If percentChange > greatestIncrease Then
                    greatestIncrease = percentChange
                    giTicker = ws.Cells(n, 1)
                ElseIf percentChange < greatestDecrease Then
                    greatestDecrease = percentChange
                    gdTicker = ws.Cells(n, 1)
                End If
            totalVolume = totalVolume + ws.Cells(n, 7)
                If totalVolume > greatestTotalVolume Then
                    greatestTotalVolume = totalVolume
                    gtvTicker = ws.Cells(n, 1)
                End If
                
            ws.Cells(nextRow, 9) = ws.Cells(n, 1)
            ws.Cells(nextRow, 10) = totalVolume
            openedAt = ws.Cells(n + 1, 3)
            yearlyChange = 0
            totalVolume = 0
            nextRow = nextRow + 1
        Else
            totalVolume = totalVolume + ws.Cells(n, 7)
        End If
    Next n
    ws.Cells(1, 15) = "Ticker"
    ws.Cells(1, 16) = "Value"
    ws.Cells(2, 14) = "Greatest % Increase"
    ws.Cells(3, 14) = "Greatest % Decrease"
    ws.Cells(4, 14) = "Greatest Total Volume"
    ws.Cells(4, 15) = gtvTicker
    ws.Cells(2, 15) = giTicker
    ws.Cells(3, 15) = gdTicker
    ws.Cells(4, 16) = greatestTotalVolume
    ws.Cells(2, 16) = greatestIncrease
    ws.Cells(3, 16) = greatestDecrease
    ws.Cells(2, 16).NumberFormat = "0.00%"
    ws.Cells(3, 16).NumberFormat = "0.00%"

Next ws
    
End Sub