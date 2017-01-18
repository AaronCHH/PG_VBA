Attribute VB_Name = "Module1"
Sub calc1()
  Dim a() As Double
  ReDim a(3) As Double
  a(0) = 5: a(1) = 8: a(2) = -6
  ReDim a(5) As Double
  MsgBox a(2)
End Sub

' ====== 
Sub calc2()
  Dim a() As Double
  ReDim a(3) As Double
  a(0) = 5: a(1) = 8: a(2) = -6
  ReDim Preserve a(5) As Double
  MsgBox a(2)
End Sub
