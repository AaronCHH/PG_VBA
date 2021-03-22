Attribute VB_Name = "Module1"
Option Explicit
Sub TempArray()
  Dim i As Integer, n As Integer
  Dim Tc(100) As Double, TF(100) As Double
  n = 4
  Tc(0) = 40: Tc(1) = 175: Tc(2) = 245: Tc(3) = 255: Tc(4) = 200
  For i = 0 To n
    TF(i) = 9 / 5 * Tc(i) + 32
  Next i
End Sub
