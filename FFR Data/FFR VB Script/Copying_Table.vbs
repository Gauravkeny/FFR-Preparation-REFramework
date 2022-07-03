Sub Copying_Table(Model)
'
' Macro3 Macro
'

'
    Sheets(Model).Select
    Range("B3").Select
    Selection.End(xlDown).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.Copy
End Sub