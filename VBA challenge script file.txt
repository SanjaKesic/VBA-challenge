Sub Stock():

    'Loop through all sheets
    For Each ws In Worksheets

    'Declare all variables
    Dim Stock_Name As String
    Dim Stock_Open As Double
    Dim Stock_Close As Double
    Dim Stock_Volume As Double
    Dim Stock_Change As Double
    Dim Stock_Percent As Double
    Dim i, j As Integer
    Dim min_per As Double
    Dim max_per As Double
    Dim max_vol As Double
    Dim Summary_Table_Row As Long

    'Initilize all the varibles
    Stock_Volume = 0
    Stock_Open = ws.Cells(2, 3).Value
    Stock_Close = 0
    Stock_Percent = 0
    Stock_Change = 0
    max_vol = 0
    min_per = 0
    max_per = 0
    Summary_Table_Row = 2

    'Determine the LastRow
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'Assign Headers to colums and rows to summary table
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Volume"
    
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"

    'Loop through all the ticker name
    For i = 2 To LastRow

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    Stock_Name = ws.Cells(i, 1).Value
    Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
    Stock_Close = ws.Cells(i, 6).Value
    Stock_Change = Stock_Close - Stock_Open
    Stock_Percent = ((Stock_Close - Stock_Open) / Stock_Open)


    ws.Cells(Summary_Table_Row, 9).Value = Stock_Name
    ws.Cells(Summary_Table_Row, 12).Value = Stock_Volume
    ws.Cells(Summary_Table_Row, 10).Value = Stock_Change
    ws.Cells(Summary_Table_Row, 11).Value = (Stock_Percent)

    'Add one to the summary table row
    Summary_Table_Row = Summary_Table_Row + 1

    Stock_Volume = 0

    Stock_Open = ws.Cells(i + 1, 3).Value

    Stock_Close = 0

    Stock_Change = 0

    Else

    Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value

    End If
    Next i

    LastRow_Summarytable = ws.Cells(Rows.Count, 10).End(xlUp).Row

    'Format Cells
    For i = 2 To LastRow_Summarytable
        If ws.Cells(i, 10).Value > 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 4
        Else
        ws.Cells(i, 10).Interior.ColorIndex = 3
        End If
        Next i

    ws.Range("K:K").NumberFormat = "#.00%"
    Next ws

End Sub