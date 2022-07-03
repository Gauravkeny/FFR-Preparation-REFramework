Sub Copying_Chart(Model)
'
' Macro4 Macro
'

'
    Sheets(Model).Select
    ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveChart.ChartArea.Copy
End Sub