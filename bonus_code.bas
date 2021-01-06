Attribute VB_Name = "Bonus"
Sub Bonus()

Dim ws As Worksheet

'Loop through all stocks for one year
For Each ws In Worksheets

'Create the row and column headings

ws.Range("N2").Value = "Greatest % Increase"
ws.Range("N3").Value = "Greatest % Decrease"
ws.Range("N4").Value = "Greatest Volume Total"

ws.Range("O1").Value = "Ticker"
ws.Range("P1").Value = "Value"

'Declare Variables
Dim IncreaseTicker As String
Dim DecreaseTicker As String
Dim VolumeTicker As String

Dim IncreaseValue As Variant
Dim DecreaseValue As Variant
Dim VolumeValue As Variant


'Set initial conditions
IncreaseTicker = ws.Range("I2").Value
DecreaseTicker = ws.Range("I2").Value
VolumeTicker = ws.Range("I2").Value

IncreaseValue = ws.Range("K2").Value
DecreaseValue = ws.Range("K2").Value
VolumeValue = ws.Range("L2").Value

'For loop identify max and min
Lastrow = ws.Cells(Rows.Count, "I").End(xlUp).Row


            For i = 3 To Lastrow
            
                    If IsNumeric(ws.Cells(i, "K").Value) Then
            
                            If ws.Cells(i, "K").Value > IncreaseValue Then
                                    
                                    IncreaseTicker = ws.Cells(i, "I").Value
                                    IncreaseValue = ws.Cells(i, "K").Value
                            
                            ElseIf ws.Cells(i, "K").Value < DecreaseValue Then
                                    
                                    DecreaseTicker = ws.Cells(i, "I").Value
                                    DecreaseValue = ws.Cells(i, "K").Value
                            
                            End If
                    
                    
                    End If
                    
                    
                    If ws.Cells(i, "L").Value > VolumeValue Then
                    
                            VolumeTicker = ws.Cells(i, "I").Value
                            VolumeValue = ws.Cells(i, "L").Value
                            
                    End If
            
            Next i
            
'populate sheet with results
ws.Range("O2").Value = IncreaseTicker
ws.Range("O3").Value = DecreaseTicker
ws.Range("O4").Value = VolumeTicker

ws.Range("P2").Value = IncreaseValue
ws.Range("P3").Value = DecreaseValue
ws.Range("P4").Value = VolumeValue

'formatting
ws.Range("P2:P3").NumberFormat = "0.00%"

ws.Columns("N:P").AutoFit

Next ws

End Sub

