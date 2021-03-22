Attribute VB_Name = "Module1"
' =====
Type ChemData
  Elname As String
  Symbol As String * 3
  AtomNumber As Integer
  AtomWeight As Double
End Type

' =====
Sub Periodic()
  Dim Chem(118) As ChemData
  Dim Msg As String
  Chem(1).Elname = "Actinium"
  Chem(1).Symbol = "Ac"
  Chem(1).AtomNumber = 89
  Chem(1).AtomWeight = 227.0278
  Msg = "Atomic nurnber = " & Chem(1).AtomNumber
  MsgBox Msg, , Chem(1).Elname
End Sub
