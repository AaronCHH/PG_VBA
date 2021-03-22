Attribute VB_Name = "Module1"
Sub Calc()
  Dim i As Integer, nd As Integer
  Dim Temp(1000) As Double, Dens(1000) As Double
  'fetch data from the file
  Open "./TempDensH2O.csv" For Input As #1
  nd = -1
  Do
    If EOF(1) Then Exit Do
    nd = nd + 1
    Input #1, Temp(nd), Dens(nd)
  Loop
  Close #1

  'display values on the worksheet
  Sheets("Sheet1").Select
  Range("a1").Select
  
  For i = 0 To nd
    ActiveCell.Value = Temp(i)
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = Dens(i)
    ActiveCell.Offset(1, -1).Select
  Next i

  Range("a1").Select
  
  'apply conversions
  For i = 0 To nd
    Temp(i) = Temp(i) + 273.15
    Dens(i) = Dens(i) * 1000
  Next i
  
  'create a file
  Open "./DeConv.csv" For Output As #2
  
  For i = 0 To nd
    Write #2, Temp(i), Dens(i)
  Next i
  Close #2

End Sub
