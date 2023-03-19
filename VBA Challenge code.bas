Sub Ticker1():

For Each ws In Worksheets
    
Dim Worksheet As String
       
Dim i As Long
Dim j As Long
Dim TickCount As Long
Dim Lastrow As Long
Dim Lastrow2 As Long
Dim PerChange As Double
Dim GreatIncrease As Double
Dim GreatDecrease As Double
Dim GreatTotal As Double

Worksheet = ws.Name
        
      
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("k1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
        
       
TickCount = 2
j = 2
        
  
Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
       
            
For i = 2 To Lastrow
            
             
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
ws.Cells(TickCount, 9).Value = ws.Cells(i, 1).Value
ws.Cells(TickCount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
If ws.Cells(TickCount, 10).Value < 0 Then
ws.Cells(TickCount, 10).Interior.ColorIndex = 3
                
        Else
                
ws.Cells(TickCount, 10).Interior.ColorIndex = 4
                
        End If
                    
If ws.Cells(j, 3).Value <> 0 Then
PerChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
ws.Cells(TickCount, 11).Value = Format(PerChange, "Percent")
                    
        Else
                    
ws.Cells(TickCount, 11).Value = Format(0, "Percent")
                    
        End If
                    
ws.Cells(TickCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
TickCount = TickCount + 1
j = i + 1
                
        End If
            
        Next i
            
Lastrow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row

GreatTotal = ws.Cells(2, 12).Value
GreatIncrease = ws.Cells(2, 11).Value
GreatDecrease = ws.Cells(2, 11).Value
        
For i = 2 To Lastrow2
            
If ws.Cells(i, 12).Value > GreatTotal Then
GreatTotal = ws.Cells(i, 12).Value
ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                

        End If
                
If ws.Cells(i, 11).Value > GreatIncrease Then
GreatIncrease = ws.Cells(i, 11).Value
ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
        
        End If
                
If ws.Cells(i, 11).Value < GreatDecrease Then
GreatDecrease = ws.Cells(i, 11).Value
ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                           

        End If
      
      Next i
ws.Cells(2, 17).Value = Format(GreatIncrease, "Percent")
ws.Cells(3, 17).Value = Format(GreatDecrease, "Percent")
ws.Cells(4, 17).Value = Format(GreatTotal, "Scientific")
            
     

Worksheets(Worksheet).Columns("A:Z").AutoFit
            
    Next ws
        
End Sub






