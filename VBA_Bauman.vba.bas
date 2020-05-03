Attribute VB_Name = "Module1"
Sub StockData()

'Define variables before any loop
Dim Ticker As String
Dim OpenStock As Double
Dim CloseStock As Double
Dim Volume As Double
Dim Greatest_Increase As Double
Dim GI_Ticker As String
Dim Greatest_Decrease As Double
Dim GD_Ticker As String
Dim Greatest_Volume As Double
Dim GV_Ticker As String
Dim Table_A_Row As Integer
Dim a As Integer
Dim b As Integer

'Loop through worksheets in the *.xls
Dim ws As Worksheet
For Each ws In Worksheets
ws.Activate

'Build Table A then Table B

'Set variables to initial appropriate values
Table_A_Row = 2
a = 0
b = 0

'Table A Headers
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

'Table B Headers
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"


'Table A, sorting through stock data
'update i ending for last row
For i = 2 To 1000000
    a = a + 1
    
If Cells(i, 1) = "" Then
Exit For
End If
'Accounting for stocks starting in the middle of the year, i.e. 2015 PLNT
If Cells(i, 3) = 0 Then
b = b + 1
End If

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Ticker = Cells(i, 1).Value
    
        OpenStock = Cells(i + b - (a - 1), 3).Value
        'MsgBox (OpenStock)
        
        CloseStock = Cells(i, 6).Value
        'MsgBox (CloseStock)
    
        Volume = Volume + Cells(i, 7).Value
    
        Range("I" & Table_A_Row).Value = Ticker
        Range("J" & Table_A_Row).Value = CloseStock - OpenStock
        'If OpenStock = 0 to resolve Div by Zero error, now need to find a way for the counter on PLNT stock to begin after zeros
        If OpenStock = 0 Then
        Range("K" & Table_A_Row).Value = 0
        Range("K" & Table_A_Row).Interior.ColorIndex = 3
        Else
        Range("K" & Table_A_Row).Value = (CloseStock - OpenStock) / OpenStock
        End If
        Range("L" & Table_A_Row).Value = Volume

        Table_A_Row = Table_A_Row + 1
        Volume = 0
        a = 0
    Else
        Volume = Volume + Cells(i, 7).Value
    
    End If

Next i


'Conditional Formating Table A; Yearly Change; red, negative, green, positive, blank, no change
'i arbitrarily high number, loops actually ends on blank (last value)
'Table B, sorting through Table A for greatest percent increases and decreases and total volume
'Table B build incorporated in same loop as Conditional Formatting
For i = 2 To 1000000

If Cells(i, 11) = "" Then
Exit For
End If

Cells(i, 11).Value = Format(Cells(i, 11).Value, "Percent")

'Table B build
If i = 2 Then
Greatest_Increase = Cells(i, 11).Value
Greatest_Decrease = Cells(i, 11).Value
Greatest_Volume = Cells(i, 12).Value
End If

If Cells(i, 11).Value >= Greatest_Increase Then
Greatest_Increase = Cells(i, 11).Value
GI_Ticker = Cells(i, 9).Value
End If
If Cells(i, 11).Value <= Greatest_Decrease Then
Greatest_Decrease = Cells(i, 11).Value
GD_Ticker = Cells(i, 9).Value
End If
If Cells(i, 12).Value >= Greatest_Volume Then
Greatest_Volume = Cells(i, 12).Value
GV_Ticker = Cells(i, 9).Value
End If

If Cells(i, 10).Value < 0 Then
'Turn cell red
Cells(i, 10).Interior.ColorIndex = 3
ElseIf Cells(i, 10).Value > 0 Then
'Turn cell green
Cells(i, 10).Interior.ColorIndex = 4
End If
Next i

Range("P2").Value = GI_Ticker
Range("P3").Value = GD_Ticker
Range("P4").Value = GV_Ticker
Range("Q2").Value = Greatest_Increase
Range("Q2").Value = Format(Range("Q2").Value, "Percent")
Range("Q3").Value = Greatest_Decrease
Range("Q3").Value = Format(Range("Q3").Value, "Percent")
Range("Q4").Value = Greatest_Volume

Next ws


End Sub


