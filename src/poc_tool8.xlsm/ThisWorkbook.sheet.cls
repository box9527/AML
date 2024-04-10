Option Explicit

Private Sub Workbook_Open()

    Dim lastRow    As Long
    Dim wsShInData As Worksheet
    Dim wsShSimple As Worksheet
    Dim wsShMoney  As Worksheet
    Dim wsShLabel  As Worksheet

    On Error Resume Next
    Set wsShInData = Worksheets(SheetNameInputData)
    Set wsShSimple = Worksheets(SheetNameSimple)
    Set wsShMoney = Worksheets(SheetNameMoney)
    Set wsShLabel = Worksheets(SheetNameLabel)
    On Error GoTo 0

    Dim emptyData As Boolean
    lastRow = wsShInData.Cells(wsShLabel.Rows.Count, 1).End(xlUp).row
    If lastRow <= RowDataBegin Then
        emptyData = True
    End If

    If Not wsShSimple Is Nothing Then
        With Worksheets(SheetNameSimple)
            Dim arrTW    As Variant
            Dim arrOccur As Variant
            Dim arrOppon As Variant

            arrTW = ExcelUtils.GetUiTWArray()
            arrOccur = ExcelUtils.GetUiOccurArray()
            arrOppon = ExcelUtils.GetUiOpponArray()

            Dim i As Long
            For i = LBound(arrTW) To UBound(arrTW)
                .cbWindow.AddItem arrTW(i)
            Next i
            For i = LBound(arrOccur) To UBound(arrOccur)
                .cbOccurrence.AddItem arrOccur(i)
            Next i
            For i = LBound(arrOppon) To UBound(arrOppon)
                .cbOpponent.AddItem arrOppon(i)
            Next i

            If Not wsShLabel Is Nothing Then
                lastRow = wsShLabel.Cells(wsShLabel.Rows.Count, 1).End(xlUp).row
                .cbPattern.AddItem UiPatternString

                For i = 2 To lastRow
                    .cbPattern.AddItem wsShLabel.Cells(i, 1).value
                Next i
            End If
        End With
    End If

    If (Not wsShMoney Is Nothing) And (emptyData = False) Then
        Dim pt As PivotTable
        Dim rowPivot As Long

        On Error Resume Next
        Set pt = wsShMoney.PivotTables(PivotTableName03)
        rowPivot = pt.TableRange1.Cells(1, 1).row
        Set HeaderPivotTable3 = wsShMoney.Range("A" & CStr(rowPivot - 1))

        Set pt = wsShMoney.PivotTables(PivotTableName04)
        rowPivot = pt.TableRange1.Cells(1, 1).row
        Set HeaderPivotTable4 = wsShMoney.Range("H" & CStr(rowPivot - 1))

        Set pt = wsShMoney.PivotTables(PivotTableName05)
        rowPivot = pt.TableRange1.Cells(1, 1).row
        Set HeaderPivotTable5 = wsShMoney.Range("A" & CStr(rowPivot - 1))

        Set pt = wsShMoney.PivotTables(PivotTableName06)
        rowPivot = pt.TableRange1.Cells(1, 1).row
        Set HeaderPivotTable6 = wsShMoney.Range("H" & CStr(rowPivot - 1))
        On Error GoTo 0
    End If

    If emptyData = True Then
        ExcelUtils.ClearSheet SheetNameMoney
    End If

    menu.createMenu
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    On Error Resume Next
    Dim wsShOrg As Worksheet
    Set wsShOrg = Worksheets(SheetNameOrginal)
    On Error GoTo 0

    If Sh.name = SheetNameSimple Then
        If (Len(MainAccName) <= 0) And (Not wsShOrg Is Nothing) Then
            With Worksheets(SheetNameOrginal)
                MainAccName = .Cells(4, 2).value
                MainAccId = .Cells(5, 2).value
            End With
        Else
            Exit Sub
        End If

        ExcelUtils.RenewAccNameID MainAccName, MainAccId
    End If
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    menu.deleteMenu
End Sub

' 只要原始資料那頁的內容有改，就會自動執行
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    If Sh.name = SheetNameOrginal Then
        Application.ErrorCheckingOptions.NumberAsText = False
        Application.ScreenUpdating = False ' For performance
        Application.DisplayAlerts = False

        With Worksheets(SheetNameMain)
            Call .BtnAnalyze_Click
            ExcelUtils.RenewAccNameID MainAccName, MainAccId
        End With

        Application.DisplayAlerts = True
        Application.ScreenUpdating = True ' For performance
        Application.ErrorCheckingOptions.NumberAsText = True
    End If
End Sub
