Attribute VB_Name = "ExcelUtils"

Public Function GetLastRowNumber(ByRef ws As Worksheet) As Long
    Dim lastRow As Long
    If Not ws Is Nothing Then
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
        GetLastRowNumber = lastRow
        Exit Function
    End If

    GetLastRowNumber = lastRow
    Exit Function
End Function

Public Function GetLastColNumber(ByRef ws As Worksheet) As Long
    Dim lastCol As Long
    If Not ws Is Nothing Then
        lastRow = ExcelUtils.GetLastRowNumber(ws)
        lastCol = ws.Cells(lastRow, ws.columns.Count).End(xlToLeft).Column
        GetLastColNumber = lastCol
        Exit Function
    End If

    GetLastColNumber = lastCol
    Exit Function
End Function

Public Function GetATMChannelArray() As Variant
    x = Split(ATMChannelString, ",")
    GetATMChannelArray = x
End Function

Public Function GetCityNameArray() As Variant
    x = Split(CityNameString, ",")
    GetCityNameArray = x
End Function

Public Function GetUiTWArray() As Variant
    x = Split(UiTimeWindowString, ",")
    GetUiTWArray = x
End Function

Public Function GetUiOccurArray() As Variant
    x = Split(UiOccurrenceString, ",")
    GetUiOccurArray = x
End Function

Public Function GetUiOpponArray() As Variant
    x = Split(UiOpponentString, ",")
    GetUiOpponArray = x
End Function

Public Function GetAllChannelArray() As Variant
    ch = GetATMChannelArray()
    xml = Split(XMLChannelString, ",")

    Dim ot As Variant
    Dim otStr As String
    otStr = ColNoteChValMobile & "," & ColNoteChValOnline & "," & ColNoteChValPayment & "," & ColNoteChValSecurity & _
               "," & ColNoteChValFax & "," & ColNoteChValFEDI & "," & ColNoteChValTAX & "," & ColNoteChValIPASS & "," & ColNoteChValCrossBR

    ot = Split(otStr, ",")

    all = Split(Join(ch, ",") & "," & Join(xml, ",") & "," & Join(ot, ","), ",")

    GetAllChannelArray = all
End Function

Public Function GetShInDataExtraHeaderArray() As Variant
    Dim myHeaders As Variant
    Dim myHStr As String
    myHStr = ",,,,,,,," & ColShInDataAmountName & ",,,,,,," & ColShInDataTSMonthName & "," & ColShInDataTSSummaryName & _
             ",,," & ColShInDataBankCodeName & "," & ColShInDataTSTypeName & "," & ColShInDataATMLocName & _
             "," & ColShInDataATMCityName & "," & ColShInDataATMAreaName & "," & ColShInDataBrShowName & _
             "," & ColShInDataBranchCityName & "," & ColShInDataBranchAreaName & "," & ColShInDataTSLocName & _
             "," & ColShInDataTSChName & "," & ColShInDataTSOClockName & "," & ColShInDataVAccCShowName & _
             "," & ColShInDataVAccReasonName & "," & ColShInDataWAccCShowName & "," & ColShInDataPAccCShowName

    myHeaders = Split(myHStr, ",")
    GetShInDataExtraHeaderArray = myHeaders
End Function

Public Function GetShInDataColsToShSimple() As Variant()
    Dim arrColIDs()    As Variant
    arrColIDs = Array(1, 16, 4, 17, 18, 9, 10, 14, 13, 27)

    GetShInDataColsToShSimple = arrColIDs
End Function

Public Function GetShOrgColsForRawData() As Variant()
    Dim arrColIDs()    As Variant
    arrColIDs = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "O", "P")

    GetShOrgColsForRawData = arrColIDs
End Function

Public Function GetShInDataColsToFillRawData() As Variant()
    Dim arrColIDs()    As Variant
    arrColIDs = Array("A", "B", "C", "D", "E", "F", "G", "Q", "R", "I", "J", "K", "L", "N", "M")

    GetShInDataColsToFillRawData = arrColIDs
End Function

Public Function IsMorningTime(ByRef time As Variant) As Boolean
    Dim bIsTime As Boolean
    If time >= TimeValue(EarlyMorningBegin) And time <= TimeValue(EarlyMorningEnd) Then
        bIsTime = True
    End If

    IsMorningTime = bIsTime
End Function

Public Function IsAutoTSID(sid As String) As Boolean
    Dim bIsSID As Boolean
    If (sid = SelfServiceID) Or (sid = SelfServiceID_2) Then
        bIsSID = True
    End If

    IsAutoTSID = bIsSID
End Function

Public Function IsAutoClerk(clerk As String) As Boolean
    Dim isAuto As Boolean

    If (clerk = ColClerkChVal01) Or (clerk = ColClerkChVal02) Then
        isAuto = True
    End If

    IsAutoClerk = isAuto
End Function

' item 19, 2024/3/27, 保留手續費與利息後，以下暫時不會用到。
Public Function IsSkipInSummary(Text As String) As Boolean
    Dim bIsSkip As Boolean
    If Text = ColSummaryHandleFee Or Text = ColSummaryInterest Then
        bIsSkip = True
    End If

    IsSkipInSummary = bIsSkip
End Function

Private Sub NormalizeCellStartEnd(ByRef cellStart As String, ByRef cellEnd As String)
    If Len(cellStart) > 0 And Len(cellEnd) = 0 Then
        cellEnd = cellStart
    End If
    If Len(cellStart) = 0 And Len(cellEnd) > 0 Then
        cellStart = cellEnd
    End If
End Sub

Private Sub NormalizeCellRowCol(ByRef cellRow As Long, ByRef cellCol As Long)
    If cellRow > 0 And cellCol <= 0 Then
        cellCol = cellRow
    End If
    If cellRow <= 0 And cellCol > 0 Then
        cellRow = cellCol
    End If
End Sub

Private Sub NormalizeDataFormat(ByRef dataFormat As String)
    If Len(dataFormat) < 0 Or ((dataFormat <> DateFormat) And _
    (dataFormat <> TimeFormat) And (dataFormat <> NumberFormat) And _
    (dataFormat <> ForceStringFormat) And (dataFormat <> GeneralFormat)) Then
        dataFormat = ForceStringFormat
    End If
End Sub

Private Sub NormalizeDataHDirection(ByRef direction As Integer)
    If direction <> xlLeft Or direction <> xlRight Or direction <> xlCenter Then
        direction = xlCenter
    End If
End Sub

Private Sub NormalizeFontSize(ByRef size As Integer)
    If size > (FontSize * 4) Or size < FontSize Then
        size = FontSize
    End If
End Sub

Public Function ConvToColTSSummary(colSummary As String) As String
    Dim trans As String
    If Len(colSummary) > 0 Then
        If Left(colSummary, 1) = ColTSSummaryVal03KW Then
            trans = ColTSSummaryVal03
        ElseIf Right(colSummary, 1) = ColTSSummaryVal02KW Then
            trans = ColTSSummaryVal02
        ElseIf Left(colSummary, 1) = ColTSSummaryVal01KW Then
            trans = ColTSSummaryVal01
        ElseIf Left(colSummary, 1) = ColTSSummaryVal04KW Then
            trans = ColTSSummaryVal04
        Else
            ' item 17, 2024/3/18, 在3.1交易明細中的交易摘要欄位除了4類交易摘要之外，其餘都顯示原本的摘要內容。
            trans = colSummary 'ColTSSummaryValOt
        End If
    End If

    ConvToColTSSummary = trans
End Function

Public Function ConvToBankCode(colSerial As String) As String
    Dim bankCode As String
    If colSerial = "" Then
        bankCode = ""
        ConvToBankCode = bankCode
        Exit Function
    End If

    bankCode = Left(colSerial, 3)
    ConvToBankCode = bankCode
End Function

' 參數： Note, Store, Clerk, Summary
' 對照回原始資料 (Original) : ColShOrgChannel, ColShOrgBranchID, ColShOrgTSTeller, ColShOrgSummary
Public Function ConvToChannel(colNote As String, colStore As String, _
                              colClerk As String, colSummary As String, _
                              Optional IsBrNameExisted As Boolean = False) As String
    Dim channel As String
    Dim trimNote As String
    Dim trimStore As String
    Dim trimClerk As String
    Dim trimSummary As String

    trimNote = Trim(colNote)
    trimStore = Trim(colStore)
    trimClerk = Trim(colClerk)
    trimSummary = Trim(colSummary)
    allArr = GetAllChannelArray()

    ' 首先用 colNote 判斷
    If trimNote <> "" And channel = "" Then
        If Utils.ArrayContainsItem(allArr, trimNote) = True Then
            channel = trimNote
        End If
    End If
    ' FIXME: 轉出入帳號是不是都Ｖ打頭就是簽帳卡交易？　=IF(OR(COUNTIF(J9,"Ｖ*"),COUNTIF(K9,"Ｖ*")),"簽帳卡交易","")

    ' 再來判斷 colClerk, colStore, colSummary
    If IsAutoClerk(trimClerk) = True Then
        channel = Trim(channel & " " & ColShInDataChSAPostfix)
    ElseIf (IsAutoTSID(trimStore) = False) And (IsBrNameExisted = True) Then
        ' 自動化交易
        channel = Trim(channel & " " & ColShInDataChBRPostfix)
    ElseIf trimSummary = ColShInDataChWTPostfix Then
        channel = Trim(channel & " " & ColShInDataChWTPostfix)
    End If

    ConvToChannel = channel
End Function

Public Function ConvPivotAccShowName(vAccName As String, wAccName As String) As String
    Dim combined As String
    Dim wan As String
    Dim van As String
    wan = Trim(wAccName)
    van = Trim(vAccName)

    If (Len(wan) > 0) And (Len(van) <= 0) Then
        combined = wan
    ElseIf (Len(wan) <= 0) And (Len(van) > 0) Then
        combined = van
    ElseIf (Len(wan) > 0) And (Len(van) > 0) Then
        combined = "(" & wan & ") " & van
    Else
        combined = ""
    End If

    ConvPivotAccShowName = combined
End Function

Public Sub ClearSheet(sheetName As String, Optional clearAll As Boolean = True)
    Dim ws As Worksheet

    ' Check If the sheet exists
    On Error Resume Next
    Set ws = Worksheets(sheetName)
    On Error GoTo 0

    ' If the sheet exists, clear all data And formats
    If Not ws Is Nothing Then
        ws.Rows.Hidden = False
        ws.Cells.UnMerge
        If clearAll = True Then
            ' 一般情況都是全部清除
            ws.Cells.Clear
        Else
            ' 特別情況，為了 3.1交易明細
            ws.Rows(RowShSimpleNotEmpty).Clear
        End If

        ' 為了 3.1交易明細特別做一次漂白
        ws.Rows(RowShSimpleEmpty).Interior.color = ColorWhite
    Else
        MsgBox "Sheet        '" & sheetName & "' Not found.", vbExclamation
    End If
End Sub

Public Sub RemoveAllCharts(sheetName As String)
    Dim ws       As Worksheet
    Dim chartObj As chartObject

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

Public Sub RemoveAllPivotTables(sheetName As String)
    Dim ws       As Worksheet

    ' Check If the sheet exists
    On Error Resume Next
    Set ws = Worksheets(sheetName)
    On Error GoTo 0

    ' 刪除現有的樞紐分析表（如果存在）
    For Each PivotTable In ws.PivotTables
        PivotTable.TableRange2.Clear
    Next PivotTable
End Sub

' "Optional" is an important key. Must to have.
Public Sub ApplyFontStyle(ws As Worksheet, Optional bold As Boolean = False, _
                          Optional size As Integer = 0, Optional wrapText As Boolean = False, _
                          Optional colStart As String = "", Optional colEnd As String = "")
    If Not ws Is Nothing Then
        NormalizeCellStartEnd colStart, colEnd
        NormalizeFontSize size

        If Len(colStart) > 0 And Len(colEnd) > 0 Then
            With ws.columns(colStart & ":" & colEnd)
                .Font.name = FontName
                .Font.size = size
                .Font.bold = bold
                .wrapText = wrapText
            End With
        Else
            With ws.Cells
                .Font.name = FontName
                .Font.size = size
                .Font.bold = bold
                .wrapText = wrapText
            End With
        End If
    End If
End Sub

Public Sub ApplyRichFontStyle(ws As Worksheet, _
                          Optional cellStart As String = "", Optional cellEnd As String = "", _
                          Optional bold As Boolean = False, Optional wrapText As Boolean = False, _
                          Optional direction As Integer = xlLeft, _
                          Optional fontColor As Long = ColorBlack, Optional fontInterColor As Long = ColorWhite, _
                          Optional enableFilter As Boolean = False)
    If Not ws Is Nothing Then
        NormalizeCellStartEnd cellStart, cellEnd
        NormalizeDataHDirection direction

        If Len(cellStart) > 0 And Len(cellEnd) > 0 Then
            With ws.Range(cellStart & ":" & cellEnd)
                .Font.color = fontColor
                .Font.bold = bold
                .wrapText = wrapText
                .Interior.color = fontInterColor
                .HorizontalAlignment = direction
                If enableFilter = True Then
                    .AutoFilter
                End If
            End With
        End If
    End If
End Sub

Public Sub ApplyDataFormat(ws As Worksheet, dataFormat As String, _
                           Optional cellStart As String = "", Optional cellEnd As String = "")
    If Not ws Is Nothing Then
        NormalizeCellStartEnd cellStart, cellEnd
        NormalizeDataFormat dataFormat

        If Len(cellStart) > 0 And Len(cellEnd) > 0 Then
            With ws.columns(cellStart & ":" & cellEnd)
                .NumberFormat = dataFormat
            End With
        Else
            With ws.Cells
                .NumberFormat = dataFormat
            End With
        End If
    End If
End Sub

Public Sub ApplyDataHDirection(ws As Worksheet, direction As Integer, _
                           Optional cellStart As String = "", Optional cellEnd As String = "")
    If Not ws Is Nothing Then
        NormalizeCellStartEnd cellStart, cellEnd
        NormalizeDataHDirection direction

        If Len(cellStart) > 0 And Len(cellEnd) > 0 Then
            With ws.columns(cellStart & ":" & cellEnd)
                .HorizontalAlignment = direction
            End With
        Else
            With ws.Cells
                .HorizontalAlignment = direction
            End With
        End If
    End If
End Sub

Public Sub ApplyCellsValue(ws As Worksheet, value As Variant, _
                           Optional cellFormat As String = ForceStringFormat, _
                           Optional cellHDirection As Integer = xlCenter, _
                           Optional cellRow As Long = 0, Optional cellCol As Long = 0)
    If Not ws Is Nothing Then
        If Len(CStr(value)) > 0 Then
            NormalizeCellRowCol cellRow, cellCol
            NormalizeDataHDirection cellHDirection
            NormalizeDataFormat cellFormat

            If cellRow > 0 And cellCol > 0 Then
                With ws.Cells(cellRow, cellCol)
                    .NumberFormat = cellFormat
                    .value = value
                End With
            End If
        End If
    End If
End Sub

Public Sub ApplyRangeValue(ws As Worksheet, value As Variant, _
                           Optional cellFormat As String = ForceStringFormat, _
                           Optional cellHDirection As Integer = xlCenter, _
                           Optional cellStart As String = "", Optional cellEnd As String = "")
    If Not ws Is Nothing Then
        If Len(CStr(value)) > 0 Then
            NormalizeCellStartEnd cellStart, cellEnd
            NormalizeDataHDirection cellHDirection
            NormalizeDataFormat cellFormat

            If Len(cellStart) > 0 And Len(cellEnd) > 0 Then
                With ws.Range(cellStart & ":" & cellEnd)
                    .NumberFormat = cellFormat
                    .value = value
                End With
            End If

        End If
    End If
End Sub

' 2.2清整後資料 頁面 style
Public Sub ApplySheetInputDataStyle()
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = Worksheets(SheetNameInputData)
    On Error GoTo 0

    ExcelUtils.ApplyFontStyle ws
    ExcelUtils.ApplyDataFormat ws, DateFormat, "A", "B"
    ExcelUtils.ApplyDataFormat ws, ForceStringFormat, "C", "C"
    ExcelUtils.ApplyDataFormat ws, TimeFormat, "D", "D"
    ExcelUtils.ApplyDataFormat ws, ForceStringFormat, "E", "G"
    ExcelUtils.ApplyDataFormat ws, NumberFormat, "H", "I"
    ExcelUtils.ApplyDataHDirection ws, xlRight, "H", "I"
    ExcelUtils.ApplyDataFormat ws, ForceStringFormat, "J", "O"
    ExcelUtils.ApplyDataFormat ws, GeneralFormat, "P", "P"
    ExcelUtils.ApplyDataFormat ws, NumberFormat, "Q", "R"
    ExcelUtils.ApplyDataHDirection ws, xlRight, "Q", "R"
    ExcelUtils.ApplyDataFormat ws, ForceStringFormat, "S", "AG"
    ExcelUtils.ApplyDataHDirection ws, xlLeft, "S", "AG"

    ws.Rows(RowDataBegin - 1).HorizontalAlignment = xlCenter
    ws.Cells.columns.AutoFit
End Sub

' 3.1交易明細 頁面 style
Public Sub ApplySheetSimpleStyle()
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = Worksheets(SheetNameSimple)
    On Error GoTo 0

    ExcelUtils.ApplyFontStyle ws
    ExcelUtils.ApplyDataFormat ws, DateFormat, "A", "A"
    ExcelUtils.ApplyDataHDirection ws, xlLeft, "A", "A"
    ExcelUtils.ApplyDataFormat ws, ForceStringFormat, "B", "B"
    'timeCol = "C"
    ExcelUtils.ApplyDataFormat ws, TimeFormat, "C", "C"
    ExcelUtils.ApplyDataHDirection ws, xlCenter, "C", "C"
    ExcelUtils.ApplyDataFormat ws, NumberFormat, "D", "F"
    ExcelUtils.ApplyDataHDirection ws, xlRight, "D", "F"
    ExcelUtils.ApplyDataFormat ws, ForceStringFormat, "G", "M"

    ExcelUtils.ApplyRichFontStyle ws, "A8", "M8", True, False, xlCenter, ColorWhite, ColorGreen, True

    ' 不要autofit，這頁手動畫好了，autofit 會跑版
    'ws.Cells.columns.AutoFit
    ' 相反的，把每個欄位都固定寬度
    For i = 1 To 13 'M, 調查結果
        If (i >= 1) And (i <= 5) Then
            ws.columns(i).ColumnWidth = ColShSimpleSmallColW
        Else
            ws.columns(i).ColumnWidth = ColShSimpleColWidth
        End If
    Next i
End Sub

' 3.2金流與交易對手 頁面 style
Public Sub ApplySheetMoneyStyle()
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = Worksheets(SheetNameMoney)
    On Error GoTo 0

    ' If the sheet exists, set style
    If Not ws Is Nothing Then
        ExcelUtils.ApplyFontStyle ws

        Dim cell As Range
        For Each cell In ws.UsedRange
            ' Original: cell.NumberFormat = "#,##0.00"
            If IsNumeric(cell.value) And cell.NumberFormat = NumberFormat Then
                cell.HorizontalAlignment = xlRight
            Else
                cell.HorizontalAlignment = xlCenter
            End If
        Next cell

        Dim usedCols As Double
        usedCols = ws.Cells(1, columns.Count).End(xlToLeft).Column
        For i = 1 To usedCols
            ws.columns(i).AutoFit
        Next i
    Else
        MsgBox "Sheet        '" & sheetName & "' Not found.", vbExclamation
    End If

    ws.Cells.columns.AutoFit
End Sub

Public Sub ActiveSheet(ByRef ws As Worksheet, Optional freezeAboveRow As Long = 0)
    If Not ws Is Nothing Then
        ws.Activate
        ActiveWindow.ScrollRow = 1
        If freezeAboveRow > 0 Then
            With ActiveWindow
                If .FreezePanes Then .FreezePanes = False
                .SplitRow = freezeAboveRow
                .FreezePanes = True
            End With
        End If
    End If
End Sub

Public Sub HighlightRow(sheetName As String, rowNum As Long, reason As String, color As Long)
    Exit Sub ' Disable the function temporaily
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = Worksheets(SheetNameSimple)
    On Error GoTo 0

    ws.Cells(rowNum, ColAlertReason).value = ws.Cells(row, ColAlertReason).value & " " & reason
    ws.Rows(rowNum).Font.color = color
End Sub

' Create bar charts in page Main
' 這個子函式用來 "取出暫存檔裡對應的值，畫到2.1 主頁面"
Public Sub CreateBarChart(rowIndex1 As Integer, rowIndex2 As Integer, _
                          targetPosition As Range, caption As String, yValue As String)
    Dim srcWs       As Worksheet
    Dim distWs      As Worksheet
    Dim chartObj    As chartObject
    Dim lastCol1    As Integer, lastCol2 As Integer
    Dim dataRange1  As Range, dataRange2 As Range
    Dim xAxisRange  As Range, yAxisRange As Range
    Dim chartRange  As Range
    ' Set the worksheet
    Set srcWs = Worksheets(SheetNameIntermediate)
    Set distWs = Worksheets(SheetNameMain)
    ' Find the last column in each row
    lastCol1 = srcWs.Cells(rowIndex1, srcWs.columns.Count).End(xlToLeft).Column
    ' Assume 跟col1 一樣
    lastCol2 = lastCol1
    ' Define the data ranges For the rows
    Set dataRange1 = srcWs.Range(srcWs.Cells(rowIndex1, 1), srcWs.Cells(rowIndex1, lastCol1))
    Set dataRange2 = srcWs.Range(srcWs.Cells(rowIndex2, 1), srcWs.Cells(rowIndex2, lastCol2))
    ' Combine the data ranges into one range
    Set chartRange = Union(dataRange1, dataRange2)
    Set xAxisRange = dataRange1
    Set yAxisRange = dataRange2
    ' Create a New chart on the worksheet
    Set chartObj = distWs.ChartObjects.Add(Left:=targetPosition.Left, Top:=targetPosition.Top, _
                                           Width:=targetPosition.Width, Height:=targetPosition.Height)
    ' Set the chart data source
    chartObj.Chart.SetSourceData chartRange
    chartObj.Chart.SeriesCollection(1).XValues = xAxisRange
    chartObj.Chart.SeriesCollection(1).values = yAxisRange
    ' Add chart title And axis labels If needed
    chartObj.Chart.HasTitle = True
    chartObj.Chart.ChartTitle.Text = caption
    chartObj.Chart.Axes(xlCategory, xlPrimary).HasTitle = False
    chartObj.Chart.Axes(xlValue, xlPrimary).HasTitle = True
    chartObj.Chart.Axes(xlValue, xlPrimary).AxisTitle.Text = yValue
    chartObj.Chart.HasLegend = False

    chartObj.Chart.ChartArea.Format.TextFrame2.TextRange.Font.name = FontName
    chartObj.Chart.ChartArea.Format.TextFrame2.TextRange.Font.size = CharAreaFontSize
    chartObj.Chart.ChartTitle.Format.TextFrame2.TextRange.Font.name = FontName
    chartObj.Chart.ChartTitle.Format.TextFrame2.TextRange.Font.size = CharTitleFontSize
End Sub
