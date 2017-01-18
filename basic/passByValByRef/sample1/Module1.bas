Attribute VB_Name = "Module1"
Sub Text()
  x = 1
  y = 0
  MsgBox "Before Call: x = " & x & " and y = " & y
  Call ValRef((x) y)
  MsgBox "Before Call: x = " & x & " and y = " & y
End Sub

Sub ValRef(x, y)
  y = x + 1
  x = 0
  MsgBox "Within Sub: x = " & x & " and y = " & y
End
