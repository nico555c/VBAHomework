Sub vbahomework():

Dim ws As Worksheet

For Each ws In Worksheets

Dim i As Double
Dim ticker As String
Dim date_open As Double
Dim date_closed As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim volume As Double
Dim open_stock As Double
Dim closing_stock As Double

Set Rng = Range("A:A")
Dim Summary_Table_Row As Double
  Summary_Table_Row = 2
Set rng_2 = Range("N:N")
Set rng_3 = Range("O:O")


ws.Cells(1, 13).Value = "Ticker"
ws.Cells(1, 14).Value = "Yearly Change"
ws.Cells(1, 15).Value = "Percent Change"
ws.Cells(1, 16).Value = "Total Stock Volume"

'count rows to define i
Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row


volume = 0
open_stock = ws.Cells(2, 3).Value
ws.Cells(2, 18) = open_stock


For i = 2 To Lastrow

 If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
 
 'find unique ticker
  ticker = ws.Cells(i, 1).Value
 
 'add volume
 volume = volume + ws.Cells(i, 7).Value
 
 ' find closing stock
 closing_stock = ws.Cells(i, 6).Value
 
'find open_stock
open_stock = ws.Cells(i + 1, 3).Value

  
 'write the ticker in the table
 ws.Range("M" & Summary_Table_Row).Value = ticker
     
 
 'write volume
 ws.Range("P" & Summary_Table_Row).Value = volume
 
 'write closing stock
 ws.Range("Q" & Summary_Table_Row).Value = closing_stock
 ws.Range("Q:Q").Font.ColorIndex = 2
 
 'write opening stock
 ws.Range("R" & Summary_Table_Row + 1).Value = open_stock
 ws.Range("R:R").Font.ColorIndex = 2
 
 ' calculate yearly change
 yearly_change = ws.Range("Q" & Summary_Table_Row).Value - ws.Range("R" & Summary_Table_Row).Value

 
 
 'write yearly range
 ws.Range("n" & Summary_Table_Row).Value = yearly_change
 
 'calculate percent change
 percent_change = ws.Range("Q" & Summary_Table_Row).Value / ws.Range("R" & Summary_Table_Row).Value - 1

  
 'write percent change
 ws.Range("o" & Summary_Table_Row).Value = percent_change
 
 ' Add one to the summary table row
    Summary_Table_Row = Summary_Table_Row + 1
 
 'set vol to zero
 
 volume = 0
 
      
 Else
 'calculate the volume
 volume = volume + ws.Cells(i, 7).Value
      
     
End If

Next i
'formating
ws.Range("N:N").FormatConditions.Delete
'red
ws.Range("N:N").FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
ws.Range("N:N").FormatConditions(1).Interior.Color = RGB(255, 0, 0)

'green
ws.Range("N:N").FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
ws.Range("N:N").FormatConditions(2).Interior.Color = RGB(0, 128, 0)


'percentage format
ws.Range("O:O").NumberFormat = "0.00%"

ws.Cells(1, 14).FormatConditions.Delete

'Range(rng_3).Value = FormatPercent(Range(rng_3))
'rng_2.FormatConditions.AddDatabar

Next ws

End Sub



