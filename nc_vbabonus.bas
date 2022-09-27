Sub vbabonus():

Dim ws As Worksheet

For Each ws In Worksheets

Dim i As Double
Dim ticker As String
Dim max_increase As Double
Dim decrease As Double
Dim g_volume As Double
Dim Lastrow As Double
  
ws.Cells(1, 22).Value = "Ticker"
ws.Cells(1, 23).Value = "Value"

ws.Cells(2, 21).Value = "Greates % Increase"
ws.Cells(3, 21).Value = "Greates % Decrease"
ws.Cells(4, 21).Value = "Greates Volume"
  
Lastrow = ws.Cells(Rows.Count, 14).End(xlUp).Row

' max increase
max_increase = 0

For i = 2 To Lastrow

 If ws.Cells(i, 15).Value >= max_increase Then
 max_increase = ws.Cells(i, 15).Value
 ticker = ws.Cells(i, 13).Value

 
 
 End If

Next i
 ws.Cells(2, 22).Value = ticker
 ws.Cells(2, 23).Value = max_increase
 ws.Cells(2, 23).NumberFormat = "0.00%"

'max decrease
decrease = 0
For i = 2 To Lastrow

 If ws.Cells(i, 15).Value < decrease Then
 decrease = ws.Cells(i, 15).Value
 ticker = ws.Cells(i, 13).Value

 
 
 End If

Next i
 ws.Cells(3, 22).Value = ticker
 ws.Cells(3, 23).Value = decrease
 ws.Cells(3, 23).NumberFormat = "0.00%"


'total volume
g_volume = 0
For i = 2 To Lastrow

 If ws.Cells(i, 16).Value > g_volume Then
 g_volume = ws.Cells(i, 16).Value
 ticker = ws.Cells(i, 13).Value

 
 
 End If

Next i
 ws.Cells(4, 22).Value = ticker
 ws.Cells(4, 23).Value = g_volume
 
 
Next ws

End Sub
