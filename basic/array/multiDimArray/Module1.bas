Attribute VB_Name = "Module1"
Sub Test()
  
  Dim nr As Integer, nc As Integer
  Dim Tc(3, 3) As Double, Tf(3, 3) As Double
  
  nr = 3
  nc = 3

  Tc(0, 0) = 70: Tc(0, 1) = 60: Tc(0, 2) = 50: Tc(0, 3) = 30
  Tc(1, 0) = 80: Tc(1, 1) = 70: Tc(1, 2) = 60: Tc(1, 3) = 50
  Tc(2, 0) = 90: Tc(2, 1) = 80: Tc(2, 2) = 70: Tc(2, 3) = 60
  Tc(3, 0) = 95: Tc(3, 1) = 90: Tc(3, 2) = 80: Tc(3, 3) = 70

  For i = 0 To nr
    For j = 0 To nc
      Tf(i, j) = 9 / 5 * Tc(i, j) + 32
      Debug.Print (Tf(i, j))
    Next
  Next
End Sub
