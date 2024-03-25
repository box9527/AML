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
            .cbPattern.AddItem ws.Cells(i, 1).value
        Next i
    End With

    Dim pt As PivotTable
    Dim rowPivot As Long

    Set ws = ThisWorkbook.Sheets(SheetNameMoney)

    If Not ws Is Nothing Then
        On Error Resume Next
        Set pt = ws.PivotTables("MyPivotTable3")
        rowPivot = pt.TableRange1.Cells(1, 1).row
        Set HeaderPivotTable3 = ws.Range("A" & CStr(rowPivot - 1))

        Set pt = ws.PivotTables("MyPivotTable4")
        rowPivot = pt.TableRange1.Cells(1, 1).row
        Set HeaderPivotTable4 = ws.Range("H" & CStr(rowPivot - 1))

        Set pt = ws.PivotTables("MyPivotTable5")
        rowPivot = pt.TableRange1.Cells(1, 1).row
        Set HeaderPivotTable5 = ws.Range("A" & CStr(rowPivot - 1))

        Set pt = ws.PivotTables("MyPivotTable6")
        rowPivot = pt.TableRange1.Cells(1, 1).row
        Set HeaderPivotTable6 = ws.Range("H" & CStr(rowPivot - 1))


        On Error GoTo 0

    End If

    menu.createMenu
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    If Sh.name = SheetNameSimple Then
        Sh.Range("O1").value = MainAccName
        Sh.Range("O2").value = MainAccId
    End If
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    menu.deleteMenu
End Sub