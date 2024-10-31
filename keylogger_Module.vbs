Option Explicit

'================================================================
'Develop PERSONAL.xls
'[module] reload Macro
'================================================================

Sub reloadPERSONALxls()

    'ThisWorkbook.Save
    Application.OnTime Now, "OpenBook"
    ThisWorkbook.Close

End Sub
Private Sub OpenBook()
    ThisWorkbook.Activate
End Sub
