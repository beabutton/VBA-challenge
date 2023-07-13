VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stonk()
    Dim ws As Worksheet
    For Each ws In Worksheets
' add columns
    ws.Cells(1, 9).Value = "Ticker Symbol"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
'tick return
    Dim tick As String
        tick = " "
    Dim tickrow As Double
        tickrow = 1
    Dim i As Double
    Dim endwb As Double
        endwb = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To endwb
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                tickrow = tickrow + 1
                tick = ws.Cells(i, 1).Value
                ws.Cells(tickrow, 9).Value = tick
            End If
        Next i
             
' open value
    Dim jan As Double
    Dim orow As Double
        orow = 1
    Dim j As Double

        For j = 2 To endwb
            If Right(ws.Cells(j, 2).Value, 4) = "0102" Then
                orow = orow + 1
                jan = ws.Cells(j, 3).Value
                ws.Cells(orow, 25).Value = jan
            End If
        Next j
        
' close value
    Dim dec As Double
    Dim crow As Double
        crow = 1
    Dim k As Double
        
        For k = 2 To endwb
            If Right(ws.Cells(k, 2).Value, 4) = "1231" Then
                crow = crow + 1
                dec = ws.Cells(k, 6).Value
                ws.Cells(crow, 26).Value = dec
            End If
        Next k
        
' yearly change
    Dim l As Double
    Dim change As Double
        For l = 2 To endwb
            If ws.Cells(l, 25) <> 0 Then
                ws.Cells(l, 10).Value = ws.Cells(l, 26).Value - ws.Cells(l, 25).Value
            End If
            If ws.Cells(l, 10).Value >= 0# Then
                ws.Cells(l, 10).Interior.ColorIndex = 4
            ElseIf ws.Cells(l, 10).Value < -0 Then
                ws.Cells(l, 10).Interior.ColorIndex = 3
            End If
        Next l
      
' % change
    Dim m As Double
    Dim perchange As Double
        For m = 2 To endwb
            ws.Range("K" & m).NumberFormat = "0.00%"
            If ws.Cells(m, 10).Value <> 0# Then
                ws.Cells(m, 11).Value = (ws.Cells(m, 26).Value - ws.Cells(m, 25).Value) / ws.Cells(m, 25).Value
            End If
        Next m
        
' total volume
    Dim n As Double
    Dim vol As Double
        vol = 0
    Dim stockrow As Double
        stockrow = 2

        For n = 2 To endwb
            If ws.Cells(n, 1).Value <> ws.Cells(n + 1, 1).Value Then
                vol = vol + ws.Cells(n, 7).Value
                ws.Cells(stockrow, 12).Value = vol
                stockrow = stockrow + 1
                vol = 0
                Else
                    vol = vol + ws.Cells(n, 7).Value
            End If
        Next n

' greatest % increase
    Dim endsum As Double
        endsum = 3001
    For n = 2 To endsum
        If ws.Cells(n, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & endsum)) Then
            ws.Cells(2, 17).Value = ws.Cells(n, 11).Value
            ws.Cells(2, 17).NumberFormat = "0.00%"
        End If
    Next n
    
' greatest % decrease
    For n = 2 To endsum
        If ws.Cells(n, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & endsum)) Then
            ws.Cells(3, 17).Value = ws.Cells(n, 11).Value
            ws.Cells(3, 17).NumberFormat = "0.00%"
        End If
    Next n
    
' total stock volume
    For n = 2 To endsum
        If ws.Cells(n, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & endsum)) Then
            ws.Cells(4, 17).Value = ws.Cells(n, 12).Value
        End If
    Next n

'tick summary
    For n = 2 To endsum
        If ws.Cells(2, 17).Value = ws.Cells(n, 11).Value Then
            ws.Cells(2, 16).Value = ws.Cells(n, 9).Value
        End If
            
        If ws.Cells(3, 17).Value = ws.Cells(n, 11).Value Then
           ws.Cells(3, 16).Value = ws.Cells(n, 9).Value
        End If
        
        If ws.Cells(4, 17).Value = ws.Cells(n, 12).Value Then
            ws.Cells(4, 16).Value = ws.Cells(n, 9).Value
        End If
    Next n
    
    ws.Columns.AutoFit
Next ws
End Sub








