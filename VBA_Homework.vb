
Sub multiple_year_stock_data()

For Each ws In Worksheets

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'get last row data
Dim ticker As String
Dim yearly As Double
Dim percent As Double
Dim vol As LongLong
Dim OV As Double
Dim CV As Double
Dim i As Long
Dim j As Long


'set initial values
vol = 0

 j = 2
y = 2
    
 'set column headers
 ws.Range("I1").Value = "Ticker"
 ws.Range("J1").Value = "Yearly Change"
 ws.Range("K1").Value = "Percent Change"
 ws.Range("L1").Value = "Total Volume"
 
 'loop for current ws to last row
 For i = 2 To lastrow
 
 If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
 'if cells(263,1) <> cells(262,1) then
 'if cells(524,1) <> cells (523,1) then
 'row 262
 
    CV = ws.Cells(i, 6).Value
    OV = ws.Cells(y, 3).Value
    'for the first time ov will be cells(2,3) which is 50.nnjk
    ' for the second time y is at 263
    'cv= cells(523,6)
    'ov =cells(263,3)
    
    yearly = CV - OV
    If OV = 0 Then
        percent = 0
    Else
        percent = Round(yearly / OV * 100, 2)
    End If

    vol = vol + ws.Cells(i, 7).Value
    
     'Print the results
    ws.Range("i" & j) = ws.Cells(i, 1).Value
    ws.Range("j" & j) = yearly
    ws.Range("k" & j) = percent & "%"
    ws.Range("l" & j) = vol
    
    If yearly > 0 Then
        ws.Range("j" & j).Interior.ColorIndex = 4
    Else
        ws.Range("j" & j).Interior.ColorIndex = 3
    End If
 
    
    
    'start next ticker symbol
    j = j + 1
    vol = 0
    y = i + 1
    ' so now y = 262+1
    

Else
    vol = vol + ws.Cells(i, 7).Value

    End If
    
Next i

'lastrow = ws.Cells(Rows.Count, 10).End(xlUp).Row
'For i = 2 To lastrow
'If OV <> 0 Then


    Next ws

End Sub