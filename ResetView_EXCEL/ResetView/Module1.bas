Attribute VB_Name = "Module1"
Sub ResetView()

  Dim WS_Count As Integer
  ' Dim I As Integer
  Dim Current As Worksheet

  ' Loop through all of the worksheets in the active workbook.
  For Each Current In Worksheets

    Current.Select
    ActiveWindow.Zoom = 125
    ' Current.Range("A1").End(xlDown).Select
    Cells(1, 1).Select
    ' Debug.Print ActiveSheet.Cells(1, 1).Value
    Current.Cells.Font.Name = "·L³n¥¿¶ÂÅé"
    
    ' Call CalcRlt

  Next
  
  ActiveWorkbook.Worksheets(1).Select
End Sub


