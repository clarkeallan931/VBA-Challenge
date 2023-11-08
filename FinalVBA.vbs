VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stockprices()

Dim wsCount As Integer
Dim x As Integer
wsCount = ActiveWorkbook.Worksheets.Count

Dim SummaryRow As Long
Dim Difference As Long
Dim Change As Long
Dim LR As Long
Dim result As String
Dim I As Long



For x = 1 To (wsCount + 1)
    Debug.Print x
    Sheets(x).Activate
     
SummaryRow = 2
StockCount = 0
StockVolume = 0
LR = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row
Range("H1:T1").Font.Bold = True
Range("J1").Value = "Ticker"
Range("K1").Value = "Yearly Change"
Cells(1, 12).Value = "Percent Change"
Cells(1, 13).Value = "Total Stock Volume"
Cells(1, 18).Value = "Ticker"
Cells(1, 19).Value = "Value"
Cells(2, 17).Value = "Greatest Increase"
Cells(3, 17).Value = "Greatest Decrease"
Cells(4, 17).Value = "Greatest Total Volume"


    For I = 2 To LR
       StockCount = StockCount + 1
       StockVolume = StockVolume + Cells(I, 7).Value
       If Cells(I, 1).Value <> Cells(I + 1, 1).Value Then
        Cells(SummaryRow, 10).Value = Cells(I, 1).Value
        Change = (I + 1) - StockCount
        Cells(SummaryRow, 11) = Cells(I, 6).Value - Cells(Change, 3).Value
        Cells(SummaryRow, 12) = FormatPercent((Cells(I, 6).Value - Cells(Change, 3).Value) / Cells(Change, 3).Value)
        Cells(SummaryRow, 13) = StockVolume
        StockVolume = 0
        StockCount = 0
        SummaryRow = SummaryRow + 1
        
        
        End If

        

Next I

result1 = WorksheetFunction.Max(Range("L:L"))
Range("S2").Value = FormatPercent(result1)

result2 = WorksheetFunction.Min(Range("L:L"))
Range("S3").Value = FormatPercent(result2)

result3 = WorksheetFunction.Max(Range("M:M"))
Range("S4").Value = result3


For I = 2 To 3002
  For j = 10 To 14
        If Cells(I, 12).Value = Range("S2").Value Then
    Range("R2").Value = Cells(I, 10).Value
    ElseIf Cells(I, 12).Value = Range("S3").Value Then
    Range("R3").Value = Cells(I, 10).Value
     ElseIf Cells(I, 13).Value = Range("S4").Value Then
    Range("R4").Value = Cells(I, 10).Value
    End If
    Next j
    Next I


For m = 1 To 3002

If Cells(m, 11) < 0 Then
Cells(m, 11).Interior.Color = vbRed
      
ElseIf Cells(m, 11) > 0 Then
Cells(m, 11).Interior.Color = vbGreen
      
End If

Next m

Next x

End Sub




