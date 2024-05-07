VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub stocks():

Dim ticker As String
Dim ticker_close As Double
Dim ticker_open As Double
Dim price_change As Double
Dim total As LongLong
Dim p As Long
Dim vol As LongLong
Dim i As Long
Dim percent_change As Double

For Each ws In ThisWorkbook.Worksheets
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Quarterly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    
total = 0
p = 2
ticker_open = ws.Cells(2, 3).Value
percent_change = ws.Cells(2, 10).Value

For i = 2 To 93000
    ticker = ws.Cells(i, 1).Value
    vol = ws.Cells(i, 7).Value
    
         
     If (ws.Cells(i + 1, 1).Value <> ticker) Then
        total = total + vol
        
        'Took prof help from Monday(4/29) class...
        
        ticker_close = ws.Cells(i, 6).Value
        price_change = ticker_close - ticker_open
        percent_change = price_change / ticker_open
        
        ws.Cells(p, 9).Value = ticker
        ws.Cells(p, 12).Value = total
        ws.Cells(p, 10).Value = price_change
        ws.Cells(p, 11).Value = percent_change
        ' Took help from Chritopher Madden(TA)
        Select Case price_change
            Case Is > 0
                ws.Cells(p, 10).Interior.ColorIndex = 4
            Case Is < 0
                ws.Cells(p, 10).Interior.ColorIndex = 3
            Case Else
                ws.Cells(p, 10).Interior.ColorIndex = 0
        End Select
        
        Select Case percent_change
            Case Is > 0
                ws.Cells(p, 11).Interior.ColorIndex = 4
            Case Is < 0
                ws.Cells(p, 11).Interior.ColorIndex = 3
            Case Else
                ws.Cells(p, 11).Interior.ColorIndex = 0
        End Select
            
                     
        total = 0
        p = p + 1
        ticker_open = ws.Cells(i + 1, 3).Value
    Else
        total = total + vol
            
    End If
Next i
Next ws

End Sub
        
        
             
Sub stock_increase():

Dim rng As Range
Dim highest, Valuetemp  As Double
Dim lowest As Double
Dim total_vol, Valuetemp1 As LongLong
Dim i As Long
Dim HighestTicker, LowestTicker, HighestVolticker As String
Dim lastrow As Long


For Each ws In ThisWorkbook.Worksheets
lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row
'Took help from (TA) Christopher Madden
'Took Xpert learning Assistance help.....
ws.Cells(2, 16).Value = "%" & WorksheetFunction.Max(ws.Range("K2:K" & lastrow)) * 100
ws.Cells(3, 16).Value = "%" & WorksheetFunction.Min(ws.Range("K2:K" & lastrow)) * 100
ws.Cells(4, 16).Value = WorksheetFunction.Max(ws.Range("L2:L" & lastrow))
highest = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & lastrow)), ws.Range("K2:K" & lastrow), 0)
lowest = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & lastrow)), ws.Range("K2:K" & lastrow), 0)
total_vol = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & lastrow)), ws.Range("L2:L" & lastrow), 0)
ws.Range("O2") = ws.Cells(highest + 1, 9)
ws.Range("O3") = ws.Cells(lowest + 1, 9)
ws.Range("O4") = ws.Cells(total_vol + 1, 9)
   'Took help from class room excercises.....

ws.Range("J2:J" & lastrow).NumberFormat = "0.00"
ws.Range("K2:K" & lastrow).NumberFormat = "0.00%"
ws.Range("A:P").Columns.AutoFit
Next ws
End Sub


   Sub reset()
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Range("I:Q").Value = ""
        ws.Range("I:L").Interior.ColorIndex = 2
    Next ws
End Sub
    







