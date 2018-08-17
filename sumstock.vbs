Sub sumalphastocks()

Dim Stock_Name As String
Dim Stock_Vol As Double
Dim Summary_Table_Row As Integer
Dim lastRow As Long


For Each ws In Worksheets

Stock_Vol = 0
Summary_Table_Row = 2
lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row


ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Total Stock Volume"

For i = 2 To lastRow

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

Stock_Name = ws.Cells(i, 1).Value
Stock_Vol = Stock_Vol + ws.Cells(i, 7).Value

ws.Range("I" & Summary_Table_Row).Value = Stock_Name
ws.Range("J" & Summary_Table_Row).Value = Stock_Vol
Summary_Table_Row = Summary_Table_Row + 1
Stock_Vol = 0

Else

Stock_Vol = Stock_Vol + ws.Cells(i, 7).Value

End If

Next i

Next ws


End Sub
