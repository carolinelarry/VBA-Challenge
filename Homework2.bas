Attribute VB_Name = "Module1"
Sub Homework2()

Dim ws As Worksheet
For Each ws In Worksheets

Dim LastRow As Long
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

Dim Volume As Double
Volume = 0

Dim OutputRow As Integer
OutputRow = 2

Dim Ticker As String

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

Dim OpenTicker As Double
OpenTicker = ws.Cells(2, 3).Value

Dim ClosingTicker As Double
Dim YearlyChange As Double
Dim PercentChange As Double

For i = 2 To LastRow
    

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        Ticker = ws.Cells(i, 1).Value
        
        ClosingTicker = ws.Cells(i, 6).Value
     
        Volume = Volume + ws.Cells(i, 7).Value
        
        YearlyChange = ClosingTicker - OpenTicker
        
        If OpenTicker = 0 Then
        
            PercentChange = Round((ClosingTicker - OpenTicker) * 100, 2)
        
        Else
        
            PercentChange = Round(((ClosingTicker - OpenTicker) / OpenTicker) * 100, 2)
        
        End If
      
        ws.Range("I" & OutputRow).Value = Ticker
        
        ws.Range("J" & OutputRow).Value = YearlyChange
       
        If YearlyChange < 0 Then
        
            ws.Range("J" & OutputRow).Interior.ColorIndex = 3
        
        ElseIf YearlyChange > 0 Then
        
            ws.Range("J" & OutputRow).Interior.ColorIndex = 4
        
        End If
        
        ws.Range("K" & OutputRow).Value = PercentChange & "%"

        ws.Range("L" & OutputRow).Value = Volume
    
        OutputRow = OutputRow + 1
    
        Volume = 0
        
        OpenTicker = ws.Cells(i + 1, 3).Value
    

    Else

        Volume = Volume + ws.Cells(i, 7).Value

    End If
        
   
Next i


Next

End Sub


