Sub Graph(OM_Value)
'
' Graph Macro
'

'
    Range("K4").Select
    Selection.End(xlToRight).Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = OM_Value
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=G7"
    
End Sub