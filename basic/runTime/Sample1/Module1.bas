Attribute VB_Name = "Module1"
Option Explicit
Sub Test()
  Dim i as Long
  Dim x as Double, dx As Double
  Dim Time2 As Single, Time1 As Single, Runtime As Single

  x = 1
  dx = 1.00001
  Time1 = Timer
  
  For i = 1 To 2000000
    x = Sin(x * dx)    
  Next i

  Time2 = Timer
  Runtime = Time2 - Time1
  MsgBox Runtime
End Sub
