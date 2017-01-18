Attribute VB_Name = "Module2"
Sub openWorkbook()
' change to current directory
  chdrive activeworkbook.path
  chdir activeworkbook.path

  Workbooks.open filename:="./Dummy.xlsx"
End Sub
