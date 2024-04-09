Attribute VB_Name = "Module1"
Sub stockchecker()

Dim opennum, closenum, ticknum As Integer
Dim stockname, largestock, smallstock, largetotstock As String
Dim totalvol, largetotvol As LongLong
Dim percent, largeper, smallper As Double
Dim ws As Worksheet
Dim wsdest As Worksheet
Dim copyRange As Range
Dim destCell As Range

Set wsdest = ThisWorkbook.Sheets("2018")


For Each ws In Worksheets
    Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    ticknum = 2
    totalvol = 0
    ws.Range("i1").Value = "Ticker"
    ws.Range("j1").Value = "Yearly Change"
    ws.Range("k1").Value = "Percent Change"
    ws.Range("l1").Value = "Total Stock Volume"
    
    For i = 2 To Lastrow
        If i = 2 Then
            opennum = ws.Cells(i, 3).Value
        End If
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            closenum = ws.Cells(i, 6).Value
            totalvol = totalvol + ws.Cells(i, 7).Value
            percent = (closenum - opennum) / opennum
            stockname = ws.Cells(i, 1).Value
            ws.Cells(ticknum, 9).Value = stockname
            ws.Cells(ticknum, 10).Value = closenum - opennum
            ws.Cells(ticknum, 11).Value = percent
            ws.Cells(ticknum, 11).NumberFormat = "0.00%"
            ws.Cells(ticknum, 12).Value = totalvol
            opennum = ws.Cells(i + 1, 3).Value
            closenum = 0
            totalvol = 0
            percent = 0
            stockname = ""
            ticknum = ticknum + 1
        Else
            totalvol = totalvol + ws.Cells(i, 7).Value
        End If
    Next i
    
    For j = 2 To (ws.Cells(Rows.Count, 9).End(xlUp).Row)
        If ws.Cells(j, 10).Value >= 0 Then
            ws.Cells(j, 10).Interior.ColorIndex = 4
            ws.Cells(j, 11).Interior.ColorIndex = 4
        Else
            ws.Cells(j, 10).Interior.ColorIndex = 3
            ws.Cells(j, 11).Interior.ColorIndex = 3
        End If
    Next j
    
'Lastrowcopy = ws.Cells(Rows.Count, 9).End(xlUp).Row
'Lastrowcopyto = wsdest.Cells(Rows.Count, 9).End(xlUp).Row + 1
'Set copyRange = ws.Range("i2", "l" & Lastrowcopy)
'Set destCell = wsdest.Range("i" & Lastrowcopyto)

'If Not ws.Name = wsdest.Name Then
'copyRange.Copy destCell
'Application.CutCopyMode = False
'End If


ws.Range("o2").Value = "Greatest % Increase"
ws.Range("o3").Value = "Greatest % Decrease"
ws.Range("o4").Value = "Greatest Total Volume"
ws.Range("p1").Value = "Ticker"
ws.Range("q1").Value = "Value"
Lastrows = ws.Cells(Rows.Count, 9).End(xlUp).Row

For k = 2 To Lastrows
    If ws.Cells(k, 11).Value > largeper Then
        largeper = ws.Cells(k, 11).Value
        largestock = ws.Cells(k, 9).Value
    End If
    If ws.Cells(k, 11).Value < smallper Then
        smallper = ws.Cells(k, 11).Value
        smallstock = ws.Cells(k, 9).Value
    End If
    If ws.Cells(k, 12).Value > largetotvol Then
        largetotvol = ws.Cells(k, 12).Value
        largetotstock = ws.Cells(k, 9).Value
    End If
Next k

ws.Range("p2").Value = largestock
ws.Range("p3").Value = smallstock
ws.Range("p4").Value = largetotstock
ws.Range("q2").Value = largeper
ws.Range("q2").NumberFormat = "0.00%"
ws.Range("q3").Value = smallper
ws.Range("q3").NumberFormat = "0.00%"
ws.Range("q4").Value = largetotvol

    
    
Next ws





End Sub
