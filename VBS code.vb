Sub WallStreet()

Dim ticker As String
Dim total_volume As Double
Dim percent_change As Double
Dim yearly_change As Double
Dim i As Integer
Dim j As Integer
Dim rowcount As Long
Dim k As Integer


i = 0
j = 0
k = 2

    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Volume"
    rowcount = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To 12000


If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
'ticker in summary table'
Cells(k, 9).Value = Cells(i, 1).Value
yearly_change_open = Cells(i, 3)


End If


total_volume = total_volume + Cells(i, 7).Value

Cells(k, 12).Value = total_volume


If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
    yearly_change_close = Cells(i, 6)

yearly_change = yearly_change_close - yearly_change_open

Cells(k, 10).Value = yearly_change_close - yearly_change_open


Cells(k, 11).Value = yearly_change_close / yearly_change_open - 1
total_volume = 0

k = k + 1
End If




If Cells(k, 10).Value > 0 Then
    Cells(k, 10).Interior.ColorIndex = 4
Else
    Cells(k, 10).Interior.ColorIndex = 3
    
End If



Next i




End Sub

