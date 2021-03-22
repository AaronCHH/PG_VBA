Attribute VB_Name = "Module1"
Sub Example()
  Dim i As Integer
  Dim x(4) As Double, y As Double
  
  For i = 1 To 3
    x(i) = i + 1
  Next
  
  y = 3 * x(1) + 4 * x(2) - 7 * x(3)
  msgbox y
  
End Sub
