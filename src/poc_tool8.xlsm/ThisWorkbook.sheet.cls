Option Explicit

Private Sub Workbook_Open()

    With Worksheets(SheetNameSimple)
        .cbWindow.AddItem "每日"
        .cbWindow.AddItem "每3日"
        .cbWindow.AddItem "每5日"
        .cbWindow.AddItem "每月"

        .cbOccurrence.AddItem "至少1次"
        .cbOccurrence.AddItem "至少3次"
        .cbOccurrence.AddItem "至少5次"
        .cbOccurrence.AddItem "至少10次"
        .cbOpponent.AddItem "所有人"
        .cbOpponent.AddItem "指定帳戶"
        '.cbOpponent.AddItem "任一帳戶" ' TODO:


        Dim lastRow As Long
        Dim ws As Worksheet
        Dim i As Long

        Set ws = Worksheets(SheetNameLabel)
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
        .cbPattern.AddItem "不使用"
        For i = 2 To lastRow
            .cbPattern.AddItem ws.Cells(i, 1).Value
        Next i
    End With

    menu.createMenu
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    If Sh.name = SheetNameSimple Then
        Sh.Range("O1").Value = MainAccName
        Sh.Range("O2").Value = MainAccId
    End If
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    menu.deleteMenu
End Sub