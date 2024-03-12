Option Explicit

Private Sub Workbook_Open()
    menu.createMenu
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    menu.deleteMenu
End Sub
