# TEMES GIGI SAFEQ.xlsm
Example that finds the overflow height above a weir.
```
Sub CalcH()
'
' Macro1 Macro
'
    SolverReset
    
    SolverOptions Precision:=0.001

    SolverAdd CellRef:="$B$1", Relation:=3, FormulaText:="0"

    SolverOk SetCell:="$D$5", MaxMinVal:=2, ValueOf:=0, ByChange:="$B$1", Engine:=1 _
        , EngineDesc:="GRG Nonlinear"
    
    SolverSolve UserFinish:=True, ShowRef:="ShowTrial"
    
    SolverSave SaveArea:=Range("A9")
    
End Sub

```