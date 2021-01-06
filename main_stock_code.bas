Attribute VB_Name = "Stocks"

Sub Stock()
'Declare and set worksheet
Dim ws As Worksheet

'Loop through all stocks for one year
For Each ws In Worksheets

'Create the column headings
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Opening Price"
ws.Range("K1").Value = "Closing Price"
ws.Range("L1").Value = "Yearly Change"
ws.Range("M1").Value = "Percent Change"
ws.Range("N1").Value = "Total Stock Volume"


'Define Ticker variable
Dim Ticker As String
Ticker = " "
Dim Ticker_volume As Double
Ticker_volume = 0
Dim TickerRow As Long: TickerRow = 2

'Set initial and last row for worksheet
Dim Lastrow As Long
Dim i As Long
Dim j As Integer

'Define Lastrow of worksheet
Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Fill First Row

ws.Range("I2").Value = ws.Range("A2").Value
ws.Range("J2").Value = ws.Range("C2").Value

'Do loop of current worksheet to Lastrow

For i = 3 To Lastrow

            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then 'if new stock
                
                TickerRow = TickerRow + 1
                
                'set ticker name
                Ticker = ws.Cells(i, 1).Value
                ws.Cells(TickerRow, "I").Value = Ticker
                
                'set opening value
                ws.Cells(TickerRow, "J").Value = ws.Cells(i, "C").Value
                
                'reset volume
                Ticker_volume = ws.Cells(i, "G").Value
                ws.Cells(TickerRow, "N").Value = Ticker_volume
                 
            
            Else 'if current stock
                
                'increase ticker volume
                Ticker_volume = Ticker_volume + ws.Cells(i, "G").Value
                ws.Cells(TickerRow, "N").Value = Ticker_volume
                
                'set close date
                ws.Cells(TickerRow, "K").Value = ws.Cells(i, "F").Value
                
            End If

Next i

'Calculate Yearly and Percent Changes
For r = 2 To TickerRow
    
    'yearly change
    ws.Cells(r, "L").Value = ws.Cells(r, "K").Value - ws.Cells(r, "J").Value
    
    'change colour
    If ws.Cells(r, "L").Value <= 0 Then
        ws.Cells(r, "L").Interior.ColorIndex = 3
    Else
        ws.Cells(r, "L").Interior.ColorIndex = 4
    End If
    
    'percentage change
    If ws.Cells(r, "J").Value <> 0 Then 'must check denom is not zero
        ws.Cells(r, "M").Value = ws.Cells(r, "L").Value / ws.Cells(r, "J").Value
        Else
        ws.Cells(r, "M").Value = "N/A"
    End If
    
    'format percentage
    ws.Cells(r, "M").NumberFormat = "0.00%"
    
Next r

ws.Columns("J:K").EntireColumn.Delete

ws.Columns("A:L").AutoFit

Next ws

End Sub


