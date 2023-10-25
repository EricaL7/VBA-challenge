Attribute VB_Name = "Module1"
Sub stock()

'define variables
Dim ticker As String
Dim yearly_change As Double
Dim Percentage_Change As Double
Dim open1 As Double
Dim close1 As Double
Dim Stock_Volume_Total As Variant
Dim ws As Worksheet
Dim lastrow As Long
Dim Greatest_Increase As Double
Dim Greatest_Decrease As Double
Dim Greatest_Total_Volume As Double
Dim Summary_Table_Row_Index As Integer

' Route through each Worksheet, Last Row
For Each ws In ThisWorkbook.Worksheets
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
Summary_Table_Row_Index = 2

open1 = 0
close1 = 0
Stock_Volume_Total = 0
yearly_change = 0
Percentage_Change = 0
Greatest_Increase = 0
Greatest_Decrease = 0
Greatest_Total_Volume = 0

'Column & Row Titles
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"

ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"

'i loop to populate columns
For i = 2 To lastrow
    If open1 = 0 Then
        open1 = ws.Cells(i, 3).Value
    End If
    
    Stock_Volume_Total = Stock_Volume_Total + ws.Cells(i, 7).Value
    
    'If we are at the last row for a given company
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        'Set Ticker Name
        ticker = ws.Cells(i, 1).Value
        ws.Range("I" & Summary_Table_Row_Index).Value = ticker
                    
        'Calculate Yearly Change
        close1 = ws.Cells(i, 6).Value
        yearly_change = close1 - open1
        ws.Range("J" & Summary_Table_Row_Index).Value = yearly_change
            
            'Conditional Color Formatting for Yearly Change
             If yearly_change > 0 Then
                 ws.Range("J" & Summary_Table_Row_Index).Interior.ColorIndex = 4
             Else
                 ws.Range("J" & Summary_Table_Row_Index).Interior.ColorIndex = 3
             End If
        
        'Calculate Percentage Change
        Percentage_Change = (yearly_change / open1) * 100
        ws.Range("K" & Summary_Table_Row_Index).Value = "%" & Percentage_Change
        
        'Check "Greatest Increase" Values
        If Percentage_Change > Greatest_Increase Then
            Greatest_Increase = Percentage_Change
            ws.Range("P" & 2).Value = ticker
            ws.Range("Q" & 2).Value = "%" & Greatest_Increase
        End If
        
        'Check "Greatest Decrease" Values
        If Percentage_Change < Greatest_Decrease Then
            Greatest_Decrease = Percentage_Change
            ws.Range("P" & 3).Value = ticker
            ws.Range("Q" & 3).Value = "%" & Greatest_Decrease
        End If
               
        'Calculate Total Stock Volume
        ws.Range("L" & Summary_Table_Row_Index).Value = Stock_Volume_Total
        
         'Check "Greatest Total Volume" Values
        If Stock_Volume_Total > Greatest_Total_Volume Then
            Greatest_Total_Volume = Stock_Volume_Total
            ws.Range("P" & 4).Value = ticker
            ws.Range("Q" & 4).Value = Greatest_Total_Volume
        End If
        
         
        Summary_Table_Row_Index = Summary_Table_Row_Index + 1
        open1 = 0
        Stock_Volume_Total = 0
    End If
    Next i
    ws.Range("I:Q").EntireColumn.AutoFit
    Next ws
End Sub

