Attribute VB_Name = "Module11"
Sub stocks():

For Each ws In Worksheets

Dim ticker As String
Dim start As Double
Dim finish As Double
Dim volume As Double

Dim Summary_table_row As Integer
Dim Table_Row As Integer

Application.DisplayAlerts = False
Application.ScreenUpdating = False

Table_Row = 2

volume = 0

ws.Range("I1").Value = "ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

 start = ws.Cells(2, 3).Value
 

Summary_table_row = 2


RowCount = ws.Cells(Rows.Count, 1).End(xlUp).Row



For i = 2 To RowCount

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        ticker = ws.Cells(i, 1).Value
    
        finish = ws.Cells(i, 6).Value
    
        volume = volume + Cells(i, 7).Value
      
        ws.Range("I" & Summary_table_row).Value = ticker
    
        ws.Range("L" & Summary_table_row).Value = volume
    
        yearly_change = finish - start
    
        ws.Range("J" & Summary_table_row).Value = yearly_change
        
        If start = 0 Then
        
        percent_change = yearly_change
        
        Else
        
        percent_change = (yearly_change / start)
        
        End If
        
        ws.Range("K" & Summary_table_row).Value = percent_change
        
        ws.Range("K" & Summary_table_row).NumberFormat = "0.00%"
        

    
        Summary_table_row = Summary_table_row + 1
    
        volume = 0
    
        start = ws.Cells(i + 1, 3).Value
        
        
        
        
    
    Else
        volume = volume + ws.Cells(i, 7).Value
    
    End If
    
    If Cells(i, 10).Value > 0 Then
         ws.Cells(i, 10).Interior.ColorIndex = 4
        
    ElseIf Cells(i, 10).Value < 0 Then
         ws.Cells(i, 10).Interior.ColorIndex = 3
                     
    Else
         ws.Cells(i, 10).Interior.ColorIndex = 2
        
    End If
        
        
    
Next i


Next ws


End Sub
