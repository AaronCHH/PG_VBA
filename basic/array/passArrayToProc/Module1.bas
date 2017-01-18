Attribute VB_Name = "Module1"
Sub Example()

  Dim i As Integer, nt As Integer
  Dim t(100), Temp(100) As Double
  Dim Min As Double, Max As Double
  Sheets("Sheet1").Select
  Range("a4").Select
  
  'determine the number of data in column A
  nt = ActiveCell.Row
  Selection.End(xlDown).Select
  nt = ActiveCell.Row - nt + 1
  
  'input the data
  Range("a4").Select
  
  For i = 1 To nt
    t(i) = ActiveCell.Value
    ActiveCell.Offset(0, 1).Select
    Temp(i) = ActiveCell.Value
    ActiveCell.Offset(1, -1).Select
  Next i
  
  Call MinMax(Temp, nt, Min, Max)
  
  MsgBox "minimum = " & Min & " maximum = " & Max, , _
  "Temperature"
  
End Sub


Sub MinMax(x, n, Mn, Mx)
  Dim i As Integer
  Mn = x(1)
  Mx = x(1)
  For i = 2 To n
    If x(i) < Mn Then Mn = x(i)
    If x(i) > Mx Then Mx = x(i)
  Next i
End Sub
