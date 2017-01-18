Attribute VB_Name = "Module1"
Sub adder()
  ' Steve Chapra
  ' 2/7/09
  ' input data
  Sheets("Sheet1").Select
  
  Range("b5").Select
  x = ActiveCell.Value
  
  Range("b6").Select
  y = ActiveCell.Value
  
  ' perform calculation
  z = x + y
  
  'Output results
  Range("b8").Select
  ActiveCell.Value = z

End Sub
