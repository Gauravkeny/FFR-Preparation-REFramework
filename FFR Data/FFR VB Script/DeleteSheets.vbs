Sub DeleteSheets(Models)
'
' DeleteSheets Macro
'

'
    Sheets(Models).Select
    ActiveWindow.SelectedSheets.Delete
End Sub