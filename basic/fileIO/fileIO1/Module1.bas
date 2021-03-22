Attribute VB_Name = "Module1"
Option Explicit
Sub OutputToFile()
  Dim a As Double , b As Double, StudentName As String
  a = 5
  b = 6
  StudentName = "Ima Engineer"
  Open "./test.dat" For Output As #1
  Write #1 , a , b , StudentName
  Close #1
End Sub