Attribute VB_Name = "Module1"
Sub TempSimp()
  Dim Tc0 As Double, Tc1 As Double, Tc2 As Double
  Dim Tc3 As Double, Tc4 As Double
  Dim Tf0 As Double, Tf1 As Double, Tf2 As Double
  Dim Tf3 As Double, Tf4 As Double
   
  Tc0 = 40
  Tc1 = 175
  Tc2 = 245
  Tc3 = 255
  Tc4 = 200

  Tf0 = 9 / 5 * Tc0 + 32
  Tf1 = 9 / 5 * Tc1 + 32
  Tf2 = 9 / 5 * Tc2 + 32
  Tf3 = 9 / 5 * Tc3 + 32
  Tf4 = 9 / 5 * Tc4 + 32
  
  debug.print(Tf0)
  debug.print(Tf1)
  debug.print(Tf2)
  debug.print(Tf3)
  debug.print(Tf4)
  
End Sub
