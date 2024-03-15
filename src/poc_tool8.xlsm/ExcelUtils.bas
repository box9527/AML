Attribute VB_Name = "ExcelUtils"

Public Function CheckPivotItemExisted(objPivotTable As PivotTable, strPivotFieldName As String, strPivotItemName As String) As Boolean
    '==========================
    ' �o�̥i��|�]���ѪR�X�Ӫ�����K�n���]�t"03�����b"�ӥ��ѡA
    ' �[�W�@�� For loop Check�A�ˬd�즳item�A�N���}�C
    '.PivotFields("����K�n").CurrentPage = "03�����b"
    Dim bExisted As Boolean
    For Each pivot_item In objPivotTable.PivotFields(strPivotFieldName).PivotItems
        If pivot_item.name = strPivotItemName Then
            bExisted = True
            Exit For
            Exit Function
        End If
    Next pivot_item

    Debug.Print bExisted
    bExisted = False
    Exit Function

    '==========================
End Function

Public Sub ClearSheet(sheetName As String, Optional clearAll As Boolean = True)
    Dim ws          As Worksheet
    ' Check If the sheet exists
    On Error Resume Next
    Set ws = Worksheets(sheetName)
    On Error GoTo 0
    ' If the sheet exists, clear all data And formats
    If Not ws Is Nothing Then
        ws.Rows.Hidden = False
        ws.Cells.UnMerge
        If clearAll = True Then
            ws.Cells.Clear
        Else
            ws.Rows("7:1048576").Clear
        End If

        ws.Rows("1:7").Interior.color = ColorWhite
    Else
        MsgBox "Sheet        '" & sheetName & "' Not found.", vbExclamation
    End If
End Sub

Public Sub RemoveAllCharts(sheetName As String)
    Dim ws          As Worksheet
    Dim chartObj    As chartObject
    ' Check If the sheet exists
    On Error Resume Next
    Set ws = Worksheets(sheetName)
    On Error GoTo 0
    ' If the sheet exists, remove all charts
    If Not ws Is Nothing Then
        For Each chartObj In ws.ChartObjects
            chartObj.Delete
        Next chartObj
    Else
        MsgBox "Sheet        '" & sheetName & "' Not found.", vbExclamation
    End If
End Sub

Public Sub SetSheetDefStyle(sheetName As String)
    Dim ws          As Worksheet
    On Error Resume Next
    Set ws = Worksheets(sheetName)
    On Error GoTo 0
    ' If the sheet exists, set style
    If Not ws Is Nothing Then
        With ws.Cells.Font
            .name = "�L�n������"
            .Size = 12
            '.Bold = False
        End With

        Dim cell As Range
        For Each cell In ws.UsedRange
            If IsNumeric(cell.Value) And cell.NumberFormat = "#,##0.00" Then
                cell.HorizontalAlignment = xlRight
            Else
                cell.HorizontalAlignment = xlCenter
            End If
        Next cell

        Dim usedCols As Double
        usedCols = ws.Cells(1, Columns.Count).End(xlToLeft).Column
        For i = 1 To usedCols
            ws.columns(i).AutoFit
        Next i
    Else
        MsgBox "Sheet        '" & sheetName & "' Not found.", vbExclamation
    End If    
End Sub
