Attribute VB_Name = "Module1"
Sub AccessExcel()
  Dim Answer As Double
  Sheets("Sheet1").Select
  Answer = Application.WorksheetFunction.Average(Range("A4:A24"))
  MsgBox Answer
End Sub
