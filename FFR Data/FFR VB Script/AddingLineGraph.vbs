Sub AddingLineGraph(Model)
'
' AddingLineGraph Macro
'

'
    Range("K4:M5").Select
    ActiveSheet.Shapes.AddChart2(227, xlLine).Select
    ActiveChart.SetSourceData Source:=Range(Model+"!$K$4:$M$5")
    ActiveChart.PlotArea.Select
    ActiveChart.Axes(xlValue).MajorGridlines.Select
    Selection.Delete
    ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveChart.Axes(xlValue).Select
    Selection.Delete
    ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = "FFR"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "FFR"
    With Selection.Format.TextFrame2.TextRange.Characters(1, 3).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 3).Font
        .BaselineOffset = 0
        .Bold = msoFalse
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(89, 89, 89)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 14
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Spacing = 0
        .Strike = msoNoStrike
    End With
    ActiveChart.PlotArea.Select
    ActiveChart.SetElement (msoElementDataLabelTop)
    Range("J7").Select
    ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveSheet.Shapes("Chart 1").IncrementLeft 200
    ActiveSheet.Shapes("Chart 1").IncrementTop 35
End Sub