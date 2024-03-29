' =================================================
' Sheet page 3.2金流與交易對手
' =================================================

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)

    Set ws = ThisWorkbook.Sheets(SheetNameMoney)

    ' =================================================
    ' MyPivotTable3, "轉帳存入交易By總流量"
    If Not HeaderPivotTable3 Is Nothing Then
        If Not Application.Intersect(Target, HeaderPivotTable3) Is Nothing Then
            Set pt = ws.PivotTables(PivotTableName03)
            If pt Is Nothing Then
                Exit Sub
            End If
            With pt.PivotFields(ColShMoneyAccountName)

                If .AutoSortField = ColShMoneyCountInName Then
                    .AutoSort xlDescending, ColShMoneySumIn2Name
                Else
                    .AutoSort xlDescending, ColShMoneyCountInName
                End If
            End With
        End If
    End If
    ' =================================================

    ' =================================================
    ' MyPivotTable4, 轉帳支出交易By總流量
    If Not HeaderPivotTable4 Is Nothing Then
        If Not Application.Intersect(Target, HeaderPivotTable4) Is Nothing Then
            Set pt = ws.PivotTables(PivotTableName04)
            If pt Is Nothing Then
                Exit Sub
            End If
            With pt.PivotFields(ColShMoneyAccountName)
                If .AutoSortField = ColShMoneyCountOutName Then
                    .AutoSort xlDescending, ColShMoneySumOut2Name
                Else
                    .AutoSort xlDescending, ColShMoneyCountOutName
                End If
            End With
        End If
    End If
    ' =================================================

    ' =================================================
    ' MyPivotTable5, 跨行轉帳存入交易By總流量
    If Not HeaderPivotTable5 Is Nothing Then
        If Not Application.Intersect(Target, HeaderPivotTable5) Is Nothing Then
            Set pt = ws.PivotTables(PivotTableName05)
            If pt Is Nothing Then
                Exit Sub
            End If
            With pt.PivotFields(ColShMoneyAccountName)
                If .AutoSortField = ColShMoneyCountInName Then
                    .AutoSort xlDescending, ColShMoneySumIn2Name
                Else
                    .AutoSort xlDescending, ColShMoneyCountInName
                End If
            End With
        End If
    End If
    ' =================================================

    ' =================================================
    ' MyPivotTable6, 跨行轉帳支出交易By總流量
    If Not HeaderPivotTable6 Is Nothing Then
        If Not Application.Intersect(Target, HeaderPivotTable6) Is Nothing Then
            Set pt = ws.PivotTables(PivotTableName06)
            If pt Is Nothing Then
                Exit Sub
            End If
            With pt.PivotFields(ColShMoneyAccountName)
                If .AutoSortField = ColShMoneyCountOutName Then
                    .AutoSort xlDescending, ColShMoneySumOut2Name
                Else
                    .AutoSort xlDescending, ColShMoneyCountOutName
                End If
            End With
        End If
    End If
    ' =================================================
End Sub

Private Sub Worksheet_PivotTableUpdate(ByVal Target As PivotTable)
    Exit Sub ' FIXME: to let user change filtering option. Need to reserve enough space for pivot table

    If Target.name = PivotTableName03 Or Target.name = PivotTableName04 Or _
       Target.name = PivotTableName05 Or Target.name = PivotTableName06 Then
        On Error Resume Next
        Set pivotRange = Target.TableRange1
        If Not pivotRange Is Nothing Then
            With Target.TableRange1
                .Rows.Hidden = False
                If .Rows.Count > 6 Then
                    .Rows("6:" & CStr(.Rows.Count - 1)).Hidden = True
                End If

            End With
        End If
        On Error GoTo 0
    End If
End Sub
