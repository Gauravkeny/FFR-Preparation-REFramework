Sub UpdateChart(Model,Letter)
'
' UpdateChart Macro
'

'
    ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveChart.SetSourceData Source:=Range(Model + "!$K$4:$"+Letter+"$5")
End Sub