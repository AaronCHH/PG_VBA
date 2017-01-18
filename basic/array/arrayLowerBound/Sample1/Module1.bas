Attribute VB_Name = "Module1"
Option Base 1 ' Set default array subscripts to 1.

Sub Test()
  Dim x(20) As Double, y(4, 5) As Double
  Dim t(0 To 5) As Double ' Overrides default base subscript.
  
  ' Use LBound function to test lower bounds of arrays.
  MsgBox LBound(x)      ' Returns 1.
  MsgBox LBound(y, 1)   ' Returns 1.
  MsgBox LBound(y, 2)   ' Returns 1.
  MsgBox LBound(t)      ' Returns 0.

End Sub
