Attribute VB_Name = "Module1"
Sub InefficientInput()
  Dim t(100) As Double, Temp(100) As Double
  Sheets("Sheet1").Select
  
  Range("a4").Select
  'input the data
  t(1) = ActiveCell.Value
  
  ActiveCell.Offset(0, 1).Select
  Temp(1) = ActiveCell.Value

  ActiveCell.Offset(1, -1).Select
  t(2) = ActiveCell.Value

  ActiveCell.Offset(0, 1).Select
  Temp(2) = ActiveCell.Value

  ActiveCell.Offset(1, -1).Select
  t(3) = ActiveCell.Value

  ActiveCell.Offset(0, 1).Select
  Temp(3) = ActiveCell.Value

  ActiveCell.Offset(1, -1).Select
  t(4) = ActiveCell.Value

  ActiveCell.Offset(0, 1).Select
  Temp(4) = ActiveCell.Value

  ActiveCell.Offset(1, -1).Select
  t(5) = ActiveCell.Value

  ActiveCell.Offset(0, 1).Select
  Temp(5) = ActiveCell.Value

  ActiveCell.Offset(1, -1).Select
  t(6) = ActiveCell.Value

  ActiveCell.Offset(0, 1).Select
  Temp(6) = ActiveCell.Value

  ActiveCell.Offset(1, -1).Select
  t(7) = ActiveCell.Value

  ActiveCell.Offset(0, 1).Select
  Temp(7) = ActiveCell.Value

  ActiveCell.Offset(1, -1).Select
  t(8) = ActiveCell.Value

  ActiveCell.Offset(0, 1).Select
  Temp(8) = ActiveCell.Value
  
  
  ' debug
  For i = 1 To 8
    debug.print(Temp(i))    
  Next
  

End Sub
