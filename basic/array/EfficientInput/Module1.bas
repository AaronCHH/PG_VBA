Attribute VB_Name = "Module1"
Sub InefficientInput()
  dim i as integer, nt as integer
  Dim t(100) As Double, Temp(100) As Double
    
  ' determines the number of data in column A
  Sheets("Sheet1").Select
  Range("a4").Select
  nt = ActiveCell.Row
  Selection.End(xlDown).Select
  nt = ActiveCell.Row - nt + 1
  
  'input the data
  Range("a4").Select

  For i=1 To nt
    t(i) = ActiveCell.Value  
    ActiveCell.Offset(0, 1).Select
    Temp(i) = ActiveCell.Value
    ActiveCell.Offset(1, -1).Select
  Next i
 
  ' debug
  For i = 1 To 8
    Debug.Print (Temp(i))
  Next
  
End Sub
