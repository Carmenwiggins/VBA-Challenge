Attribute VB_Name = "Module1"
Option Explicit

'Analysis for stock changes

Sub Stockanalysis()

'Declaration of variable

Dim Ticker As String
Dim Year_Change As Double
Dim Percentage_Change As Double
Dim Total_Stock_Volume As Double
Dim FinalRow As Long
Dim Open_Index As Long
Dim i As Long
Dim ws As Worksheet


 

Dim Summary_row As Long
'Dim Total_Volume As Double

For Each ws In Worksheets
ws.Activate

Summary_row = 2
Open_Index = 2
Total_Stock_Volume = 0



Range("I1").Value = "Ticker"
Range("j1").Value = "Yearly Change"
Range("K1").Value = "Percentage Change"
Range("L1").Value = "Total Stock Volume"

Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest % Total_Volume"



FinalRow = Cells(Rows.Count, "A").End(xlUp).Row

'loop through the data to find ticker for the stocks

    For i = 2 To FinalRow
    
        Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7)
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            Year_Change = (Cells(i, 6) - Cells(Open_Index, 3))
            Range("I" & Summary_row).Value = Cells(i, 1).Value
    
            Range("j" & Summary_row).Value = Year_Change
            Range("k" & Summary_row).Value = (Range("j" & Summary_row) / Cells(Open_Index, 3))
            Range("L" & Summary_row).Value = Total_Stock_Volume
            
     'to change format on PercentageChange column
     
            ws.Range("K" & Summary_row).NumberFormat = "0.00%"
            
            
     ' Use conditional formating for negative and positive values on Year change column
     
        If (Year_Change < 0) Then
            Range("j" & Summary_row).Interior.ColorIndex = 3
        ElseIf (Year_Change > 0) Then
            Range("j" & Summary_row).Interior.ColorIndex = 4
            
        End If
    
                Open_Index = i + 1
                Summary_row = Summary_row + 1
                Total_Stock_Volume = 0
    
           
        End If
  
    
 
    Next i
    
    Next ws
    
End Sub
