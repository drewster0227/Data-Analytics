Attribute VB_Name = "Module1"
Sub StockAnalysis()
    Dim ws As Worksheet
    Dim i As Double, j As Double
    Dim openPrice As Double, closePrice As Double
    Dim cellValue As String, currentStock As String
    Dim uniqueString As String
    Dim totalStocksSold As Double
    Dim greatestIncrease As Double, greatestDecrease As Double, greatestVolume As Double
    Dim greatestIncreaseStock As String, greatestDecreaseStock As String, greatestVolumeStock As String
    Dim cell As Range
    
    ' Initialize the worksheet and counters
    For Each ws In ThisWorkbook.Sheets(Array("2018", "2019", "2020"))
    openPrice = 0
    i = 2
    j = 2
    uniqueString = "Blank"
    currentStock = "Insert"
    totalStocksSold = 0
    greatestIncrease = -100 ' Initialize with a very low value
    greatestDecrease = 100 ' Initialize with a very high value
    greatestVolume = 0
    
    ' Loop through the rows until an empty cell in column A is encountered
    Do While ws.Cells(i, "A").Value <> ""
        cellValue = ws.Cells(i, "A").Value
        
        ' Check if the current stock has changed
        If cellValue <> currentStock Then
            ' Calculate and write the price change, percentage change, and total stocks sold
            If openPrice <> 0 Then
                closePrice = ws.Cells(i - 1, "F").Value
                ws.Cells(j, "J").Value = closePrice - openPrice
                ws.Cells(j, "K").Value = ((closePrice - openPrice) / openPrice)
                ws.Cells(j, "L").Value = totalStocksSold
                
                ' Check for the greatest % increase, greatest % decrease, and greatest total stock volume
                If ws.Cells(j, "K").Value > greatestIncrease Then
                    greatestIncrease = ws.Cells(j, "K").Value
                    greatestIncreaseStock = currentStock
                End If
                
                If ws.Cells(j, "K").Value < greatestDecrease Then
                    greatestDecrease = ws.Cells(j, "K").Value
                    greatestDecreaseStock = currentStock
                End If
                
                If totalStocksSold > greatestVolume Then
                    greatestVolume = totalStocksSold
                    greatestVolumeStock = currentStock
                End If
                
                j = j + 1
            End If
            ' Update the current stock, openPrice, and reset totalStocksSold
            currentStock = cellValue
            openPrice = ws.Cells(i, "C").Value
            totalStocksSold = 0 ' Reset totalStocksSold for the new stock
        End If
        
        ' Check if the unique string has changed
        If cellValue <> uniqueString Then
            uniqueString = cellValue
            ws.Cells(j, "I").Value = uniqueString
        End If
        
        ' Update totalStocksSold with the current row's volume
        totalStocksSold = totalStocksSold + ws.Cells(i, "G").Value
        
        ' Move to the next row
        i = i + 1
    Loop
    
    ' Calculate and write the values for the last stock, including total stocks sold
    If openPrice <> 0 Then
        closePrice = ws.Cells(i - 1, "F").Value
        ws.Cells(j, "J").Value = closePrice - openPrice
        ws.Cells(j, "K").Value = ((closePrice - openPrice) / openPrice)
        ws.Cells(j, "L").Value = totalStocksSold
        
        ' Check for the greatest % increase, greatest % decrease, and greatest total stock volume for the last stock
        If ws.Cells(j, "K").Value > greatestIncrease Then
            greatestIncrease = ws.Cells(j, "K").Value
            greatestIncreaseStock = currentStock
        End If
        
        If ws.Cells(j, "K").Value < greatestDecrease Then
            greatestDecrease = ws.Cells(j, "K").Value
            greatestDecreaseStock = currentStock
        End If
        
        If totalStocksSold > greatestVolume Then
            greatestVolume = totalStocksSold
            greatestVolumeStock = currentStock
        End If
    End If
   For Each cell In ws.Range("J2:J" & ws.Cells(Rows.Count, "J").End(xlUp).Row)
        ' Check if the cell value is a number
        If IsNumeric(cell.Value) Then
            ' Highlight the cell in green for positive numbers
            If cell.Value > 0 Then
                cell.Interior.ColorIndex = 4 ' 4 corresponds to green
            ' Highlight the cell in red for negative numbers
            ElseIf cell.Value < 0 Then
                cell.Interior.ColorIndex = 3 ' 3 corresponds to red
            ' Clear the background color for zero
            Else
                cell.Interior.ColorIndex = xlNone
            End If
        End If
    Next cell

    
    ' Write the greatest % increase, greatest % decrease, and greatest total stock volume to cells P2, P3, and P4
    ws.Cells(1, "I").Value = "Ticker"
    ws.Cells(1, "J").Value = "Yearly Change"
    ws.Cells(1, "K").Value = "Percent Change"
    ws.Cells(1, "L").Value = "Total Stock Volume"
    ws.Cells(2, "O").Value = "Greatest % Increase"
    ws.Cells(3, "O").Value = "Greatest % Decrease"
    ws.Cells(4, "O").Value = "Greatest Total Volume"
    ws.Cells(1, "P").Value = "Ticker"
    ws.Cells(1, "Q").Value = "Value"
    ws.Cells(2, "P").Value = greatestIncreaseStock
    ws.Cells(2, "Q").Value = greatestIncrease
    ws.Cells(3, "P").Value = greatestDecreaseStock
    ws.Cells(3, "Q").Value = greatestDecrease
    ws.Cells(4, "P").Value = greatestVolumeStock
    ws.Cells(4, "Q").Value = greatestVolume
    
    Next ws
End Sub

