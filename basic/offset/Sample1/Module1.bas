Attribute VB_Name = "Module1"
Sub Sphere()
  pii = 3.14159
  'Get value of Radius from user
  Sheets("Sheet1").Select

  Range("b5").Select
  radius = ActiveCell.Value
  
  ' Compute volume
  volume = 4 / 3 * pii * radius^3  
  ' Compute surface area
  area = 4 * pii * radius^2
  ' Display results
  Range("b7:b8").ClearContents

  Range("b7").Select
  ActiveCell.Value = area

  ActiveCell.offset(1, 0).Select
  ActiveCell.Value = volume

  Range("b5").Select
End Sub
