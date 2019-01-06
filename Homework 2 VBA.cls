VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Homework2()

Dim total As Double
    
RowCount = Cells(Rows.Count, "A").End(xlUp).Row

Range("I1").Value = "Ticker"
    
Range("J1").Value = "Total Stock Volume"
    

    
For i = 2 To RowCount
    
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
    total = total + Cells(i, 7).Value
    
    Range("I" & 2 + j).Value = Cells(i, 1).Value
     
    Range("J" & 2 + j).Value = total
    
    total = 0
        
    j = j + 1

    Else
        total = total + Cells(i, 7).Value
    

    
    End If
    
Next i
    
End Sub

