# Official doc

## SolverSolve 函數

```{vba}
Worksheets("Sheet1").Activate 
SolverReset 
SolverOptions Precision:=0.001 
SolverOK SetCell:=Range("TotalProfit"), _ 
    MaxMinVal:=1, _ 
    ByChange:=Range("C4:E6") 
SolverAdd CellRef:=Range("F4:F6"), _ 
    Relation:=1, _ 
    FormulaText:=100 
SolverAdd CellRef:=Range("C4:E6"), _ 
    Relation:=3, _ 
    FormulaText:=0 
SolverAdd CellRef:=Range("C4:E6"), _ 
    Relation:=4 
SolverSolve UserFinish:=False, ShowRef:= "ShowTrial" 
SolverSave SaveArea:=Range("A33") 
 
Function ShowTrial(Reason As Integer) 
  Msgbox Reason 
  ShowTrial = 0 
End Function
```

## References
* https://docs.microsoft.com/zh-tw/office/vba/excel/concepts/functions/solversolve-function