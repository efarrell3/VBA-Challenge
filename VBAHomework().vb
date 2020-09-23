Sub VBAHomework()

For Each ws In Worksheets

'Setting headers for summary table
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Annual Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Volume"

Dim Ticker As String
Dim StockOpen As Double
Dim StockClose As Double
Dim StockChng As Double
Dim TotalVol As Double


StockOpen = ws.Cells(2, 3).Value
StockClose = 0
StockChng = 0
TotalVol = 0

Dim Summary_Table_Row As Long
Summary_Table_Row = 2

'Counts number of rows
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row


For i = 2 To lastrow
    
    ' If the Value in the row is not equal to the next row then print that value in the box
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        'Setting ticker name
        Ticker = ws.Cells(i, 1).Value
        
        'Set closing stock price
        StockClose = ws.Cells(i, 6).Value
        
        'Add to stock volume total
        TotalVol = (TotalVol + ws.Cells(i, 7).Value)
        
        StockChng = StockClose - StockOpen
        
        'Print ticker name in summary table
        ws.Range("I" & Summary_Table_Row).Value = Ticker
        
        'Print yearly difference in summary table
        ws.Range("J" & Summary_Table_Row).Value = StockChng
            If StockChng > 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            ElseIf StockChng < 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            End If
            
        
        'Print percent change in summary table
        If StockOpen = 0 Then
        ws.Range("K" & Summary_Table_Row).Value = 0
        Else
        ws.Range("K" & Summary_Table_Row).Value = ((StockChng) / StockOpen)
        End If
        ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
        
        'Print stock volume total
        ws.Range("L" & Summary_Table_Row).Value = TotalVol
        
        'Add 1 to Summary Table Row
        Summary_Table_Row = Summary_Table_Row + 1
        
        'Set next stocks opening price
        StockOpen = ws.Cells(i + 1, 3).Value
       
       'Reset stock volume total
        TotalVol = 0
    Else
    
        TotalVol = (TotalVol + ws.Cells(i + 1, 7).Value)
        
    End If
    
    
    Next i

'Setting challenge table headers
ws.Cells(2, 14).Value = "Greatest Percent Increase"
ws.Cells(3, 14).Value = "Greatest Percent Decrease"
ws.Cells(4, 14).Value = "Greatest Total Volume"
ws.Cells(1, 15).Value = "Ticker"
ws.Cells(1, 16).Value = "Value"
ws.Cells(2, 16).NumberFormat = "0.00%"
ws.Cells(3, 16).NumberFormat = "0.00%"

'Definining challenge variables
Dim MaxNum As Double
MaxNum = 0
Dim MinNum As Double
MinNum = 0
Dim TopTot As Double
TopTot = 0
Dim Tot As Double
Tot = 0
Dim Max As Double
Max = 0
Dim Min As Double
Min = 0
Dim Ticker1 As String
Dim Ticker2 As String
Dim Ticker3 As String


'Variable that counts the number of stocks in the summary table
lastrow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row


'Greatest Percent Increase
MaxNum = WorksheetFunction.Max(ws.Columns("K"))
ws.Range("P2").Value = MaxNum

'Greatest Percent Decrease
MinNum = WorksheetFunction.Min(ws.Columns("K"))
ws.Range("P3").Value = MinNum

'Greatest Total Volume
TopTot = WorksheetFunction.Max(ws.Columns("L"))
ws.Range("P4").Value = TopTot

'Finding the ticker that corresponds to the greatest percent increase
For i = 2 To lastrow2
Max = ws.Cells(i, 11).Value
If Max = MaxNum Then
Ticker1 = ws.Cells(i, 9).Value
ws.Range("O2").Value = Ticker1
End If
Next i

'Finding the ticker that corresponds to the greatest percent decrease
For i = 2 To lastrow2
Min = ws.Cells(i, 11).Value
If Min = MinNum Then
Ticker2 = ws.Cells(i, 9).Value
ws.Range("O3").Value = Ticker2
End If
Next i

'Finding the ticker that corresponds to the greatest total volume
For i = 2 To lastrow2
Tot = ws.Cells(i, 12).Value
If Tot = TopTot Then
Ticker3 = ws.Cells(i, 9).Value
ws.Range("O4").Value = Ticker3
End If
Next i
        
    Next ws


End Sub


