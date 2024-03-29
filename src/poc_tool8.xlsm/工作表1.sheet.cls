' =================================================
' Sheet page 2.1主頁面
' =================================================

Private Const idCountAlertMonth        As Integer = 1   ' 找出 "較之前多3~5倍 超出閥值的月份 總次數", 這裡定義的是 taskID
Private Const idAmountAlertMonth       As Integer = 2   ' 找出 "較之前多3~5倍 超出閥值的月份 總金額", 這裡定義的是 taskID
Private Const idOver1MillionAlertMonth As Integer = 3   ' 找出 "當月較前月多增>=NTD100萬 超出閥值的月份", 這裡定義的是 taskID

' 這裡定義每個task 的顯示資訊的 cell 位置。都是F欄，不同列
' task 的編號，盡量與 Consts 裡的RowTask 描述相同
Private Const posRunStatus     As String = "F2"         ' 紀載在 row 2, column F
Private Const posTask1Result   As String = "F3"         ' 紀載在 row 3, column F
Private Const posTask2Result   As String = "F4"         ' 紀載在 row 4, column F
Private Const posTask3Result   As String = "F5"         ' 紀載在 row 5, column F
Private Const posTask4Result   As String = "F7"         ' 紀載在 row 7, column F
Private Const posTask5Result   As String = "F8"         ' 紀載在 row 8, column F
Private Const posTask6Result   As String = "F10"        ' 紀載在 row 10, column F
Private Const posTask9Result   As String = "F15"        ' 紀載在 row 15, column F
Private Const posTask10Result  As String = "F16"        ' 紀載在 row 16, column F
Private Const posTask11Result  As String = "F17"        ' 紀載在 row 17, column F
Private Const posTask12Result  As String = "F18"        ' 紀載在 row 18, column F
Private Const posTask13Result  As String = "F19"        ' 紀載在 row 19, column F
Private Const posTask14Result  As String = "F20"        ' 紀載在 row 20, column F
Private Const posTask7_1Result As String = "F12"        ' 紀載在 row 12, column F, 第7個任務的第1個子任務
Private Const posTask7_2Result As String = "F13"        ' 紀載在 row 12, column F, 第7個任務的第2個子任務

Private Function LookupBranchInfo(branchID As Variant) As Variant
    Dim ws          As Worksheet
    Dim lastRow     As Long
    Dim branchRange As Range
    Dim branchCell  As Range

    Set ws = ThisWorkbook.Sheets(SheetNameBranch)
    lastRow = ExcelUtils.GetLastRowNumber(ws)

    Set branchRange = ws.Range("A1:A" & lastRow)

    If Len(branchID) <= 3 Then
        branchID = Format(branchID, "0000")
    End If

    ' Search For the branch ID in column A
    Set branchCell = branchRange.Find(branchID, LookIn:=xlValues, LookAt:=xlWhole)

    ' Check If the branch ID was found
    If Not branchCell Is Nothing Then
        ' Return the corresponding value from column C
        LookupBranchInfo = branchCell.Offset(0, 3).Text & "(" & branchCell.Offset(0, 4).Text & ")"        ' Offset by 3/4olumns To Get To column D/E
        Exit Function
    End If

    ' Return an error value Or handle the Case when branch ID is Not found
    LookupBranchInfo = CVErr(xlErrValue)
End Function

Private Function FlattenResults(ByRef ws As Worksheet, _
    ByRef colIdxArray() As Long, ByRef rowAnchor As Long, Optional delimiter As String = "") As String
    Dim flattenRes As String
    If Not ws Is Nothing Then
        If Utils.IsArrayEmpty(colIdxArray) = False Then
            For i = LBound(colIdxArray) To UBound(colIdxArray)
                If colIdxArray(i) > 0 Then
                    flattenRes = flattenRes & CStr(ws.Cells(rowAnchor, colIdxArray(i)).value) & delimiter
                End If
            Next i
        End If
    End If

    FlattenResults = flattenRes
End Function

Private Function IsRatioDiffLess(ratio As Double) As Boolean
    Dim bDiffLess As Boolean
    If (ratio >= RowTask7DiffLBRatio) And (ratio <= RowTask7DiffUBRatio) Then
        bDiffLess = True
    End If

    IsRatioDiffLess = bDiffLess
End Function

Private Function FindMonthsExceedingThreshold(dateRow As Integer, amountRow As Integer, taskID As Integer) As Variant
    Dim ws              As Worksheet
    Dim lastCol         As Long
    Dim columnIndices() As Long
    Dim i               As Long

    Set ws = Worksheets(SheetNameIntermediate)
    lastCol = ExcelUtils.GetLastColNumber(ws)

    ' Initialize array To store column indices
    ReDim columnIndices(1 To 1)
    columnIndices(1) = EmptyArrayValue        ' Note:　this is a special marker for empty array

    If lastCol < 2 Then
        FindMonthsExceedingThreshold = columnIndices
        Exit Function
    End If

    Dim col                 As Long
    Dim currentMonthAmount  As Double
    Dim previousMonthAmount As Double

    i = 1
    ' Loop through the columns To identify months
    For col = 2 To lastCol        ' Assuming the first column contains headers
        ' Get the amount values For the current And previous months
        currentMonthAmount = CDbl(ws.Cells(amountRow, col).value)
        previousMonthAmount = CDbl(ws.Cells(amountRow, col - 1).value)

        ' Check If the current month's amount is three times larger than the previous month
        If taskID = idAmountAlertMonth Then
            'FIXME: 這個閥值應該參數化
            If currentMonthAmount >= 3 * previousMonthAmount Then
                ' Add the column index To the array
                ReDim Preserve columnIndices(1 To i)
                columnIndices(i) = col
                Utils.CountOne i
            End If
        ElseIf taskID = idCountAlertMonth Then
            If currentMonthAmount >= 3 * previousMonthAmount Then
                ' Add the column index To the array
                ReDim Preserve columnIndices(1 To i)
                columnIndices(i) = col
                Utils.CountOne i
            End If
        ElseIf taskID = idOver1MillionAlertMonth Then
            If currentMonthAmount >= (previousMonthAmount + 1000000) Then
                ' Add the column index To the array
                ReDim Preserve columnIndices(1 To i)
                columnIndices(i) = col
                Utils.CountOne i
            End If
        Else
            MsgBox "unknown task ID " & CStr(taskID)
        End If
    Next col

    ' Return the array of column indices
    FindMonthsExceedingThreshold = columnIndices
End Function

' TASK 1: 分月交易總金額, 總筆數
Sub ProcessRowsByMonth()
    Dim wsInput      As Worksheet
    Dim wsInter      As Worksheet
    Dim lastRow      As Long
    Dim currentGroup As Range
    Dim row          As Long
    Dim idxcolOutput As Long
    idxcolOutput = 1        ' output column index

    Dim currentYear, currentMonth As Integer

    Set wsInput = ThisWorkbook.Worksheets(SheetNameInputData)
    Set wsInter = ThisWorkbook.Worksheets(SheetNameIntermediate)
    lastRow = ExcelUtils.GetLastRowNumber(wsInput)

    ' Loop through rows starting from row rowDataBegin, day by day
    For row = RowDataBegin To lastRow
        Dim dateValue As Date
        Dim y, m      As Integer

        dateValue = wsInput.Cells(row, ColShInDataTSDate).value
        y = year(dateValue)
        m = month(dateValue)

        ' Check If the date has changed Or If it's the first row
        If row = RowDataBegin Or y <> currentYear Or m <> currentMonth Then
            ' Process the current group of rows
            If Not currentGroup Is Nothing Then
                ProcessGroupByMonth currentGroup, CStr(currentYear) & "/" & CStr(currentMonth), idxcolOutput
                Utils.CountOne idxcolOutput

                Set currentGroup = Nothing
            End If

            ' Initialize a New group For the current date
            Set currentGroup = wsInput.Range(wsInput.Cells(row, 1), wsInput.Cells(row, MaxSrcCol))
            currentYear = y
            currentMonth = m
        Else
            ' Add the row To the current group
            Set currentGroup = Union(currentGroup, wsInput.Range(wsInput.Cells(row, 1), wsInput.Cells(row, MaxSrcCol)))
        End If
    Next row

    ' Process the last group of rows
    If Not currentGroup Is Nothing Then
        ProcessGroupByMonth currentGroup, CStr(currentYear) & "/" & CStr(currentMonth), idxcolOutput
        Utils.CountOne idxcolOutput
    End If

    ' 畫圖
    ExcelUtils.CreateBarChart RowTask1TotalCountXAxis, RowTask1TotalCountYAxis, _
                              Range(RowTask1TotalCountRange), RowTask1TotalCountTitle, RowTask1TotalCountYLabel

    ExcelUtils.CreateBarChart RowTask23TotalAmountXAxis, RowTask23TotalAmountYAxis, _
                              Range(RowTask23TotalAmountRange), RowTask23TotalAmountTitle, RowTask23TotalAmountYLabel

    ' Add markers For selected one
    Dim resultArray() As Long        ' 有超過閥值的月份
    Dim i             As Integer
    Dim strSummary    As String

    ' 任務1
    With Range(posTask1Result)
        .value = ""
        resultArray = FindMonthsExceedingThreshold(RowTask1TotalCountXAxis, RowTask1TotalCountYAxis, idCountAlertMonth)
        If Utils.IsArrayEmpty(resultArray) = False Then
            strSummary = FlattenResults(wsInter, resultArray, RowTask1TotalCountXAxis, " ")

            .Font.color = ColorRed
            .wrapText = True
            .value = RowTaskAlertPrefixByMonth & strSummary & RowTaskAlertPostfixByCount
        End If
    End With

    ' 任務2
    With Range(posTask2Result)
        .value = ""
        resultArray = FindMonthsExceedingThreshold(RowTask23TotalAmountXAxis, RowTask23TotalAmountYAxis, idAmountAlertMonth)
        If Utils.IsArrayEmpty(resultArray) = False Then
            strSummary = FlattenResults(wsInter, resultArray, RowTask23TotalAmountXAxis, " ")

            .Font.color = ColorRed
            .wrapText = True
            .value = RowTaskAlertPrefixByMonth & strSummary & RowTaskAlertPostfixByAmount
        End If
    End With

    ' 任務3
    With Range(posTask3Result)
        .value = ""
        resultArray = FindMonthsExceedingThreshold(RowTask23TotalAmountXAxis, RowTask23TotalAmountYAxis, idOver1MillionAlertMonth)
        If Utils.IsArrayEmpty(resultArray) = False Then
            strSummary = FlattenResults(wsInter, resultArray, RowTask23TotalAmountXAxis, " ")

            .Font.color = ColorRed
            .wrapText = True
            .value = RowTaskAlertPrefixByMonth & strSummary & RowTaskAlertPostfixByAmount
        End If
    End With
End Sub

Sub ProcessGroupByMonth(groupRange As Range, colName As String, idxOut As Long)
    Dim sum       As Double
    Dim cell      As Range
    Dim intResult As Long
    ' groupRange represents all rows With the same date
    intResult = groupRange.row + groupRange.Rows.Count - 1

    For Each cell In groupRange.columns(ColShInDataAmtDraw).Cells
        If IsNumeric(cell.value) Then
            sum = sum + CDbl(cell.value)
        End If
    Next cell

    Dim wsInter As Worksheet
    Set wsInter = ThisWorkbook.Worksheets(SheetNameIntermediate)

    wsInter.Cells(RowTask23TotalAmountXAxis, idxOut).value = colName
    wsInter.Cells(RowTask23TotalAmountYAxis, idxOut).value = sum

    wsInter.Cells(RowTask1TotalCountXAxis, idxOut).value = colName
    wsInter.Cells(RowTask1TotalCountYAxis, idxOut).value = groupRange.Rows.Count
End Sub

' TASK
Sub ProcessRowsByDay()
    Dim wsInput                As Worksheet
    Dim wsInter             As Worksheet
    Dim lastRow           As Long
    Dim currentGroup      As Range
    Dim row               As Long
    Dim idxcolOutput      As Long
    Dim idxcolSmallAmount As Long
    Dim idxcolDiff        As Long
    idxcolOutput = 1        ' output column index
    idxcolSmallAmount = 1        ' another output column
    idxcolDiff = 1

    Dim currentYear, currentMonth, currentDay As Integer

    Set wsInput = ThisWorkbook.Worksheets(SheetNameInputData)
    Set wsInter = ThisWorkbook.Worksheets(SheetNameIntermediate)
    lastRow = ExcelUtils.GetLastRowNumber(wsInput)

    ' Loop through rows starting from row rowDataBegin, day by day
    For row = RowDataBegin To lastRow        ' lastRow FIXME: For debug only
        Dim dateValue As Date
        Dim y, m, d   As Integer
        Dim ret       As Boolean

        dateValue = wsInput.Cells(row, ColShInDataTSDate).value
        y = year(dateValue)
        m = month(dateValue)
        d = Day(dateValue)

        ' Check If the date has changed Or If it's the first row
        If row = RowDataBegin Or y <> currentYear Or m <> currentMonth Or d <> currentDay Then
            ' Process the current group of rows
            If Not currentGroup Is Nothing Then
                If ProcessGroupByDay(currentGroup, CStr(currentYear) & "/" & CStr(currentMonth) & "/" & CStr(currentDay), idxcolOutput) = True Then
                    Utils.CountOne idxcolOutput
                End If

                If ProcessGroupByDaySmallAmount(currentGroup, CStr(currentYear) & "/" & CStr(currentMonth) & "/" & CStr(currentDay), idxcolSmallAmount) = True Then
                    Utils.CountOne idxcolSmallAmount
                End If

                ProcessGroupByDayDiff currentGroup, CStr(currentYear) & "/" & CStr(currentMonth) & "/" & CStr(currentDay), idxcolDiff
                Utils.CountOne idxcolDiff

                Set currentGroup = Nothing
            End If

            ' Initialize a New group For the current date
            Set currentGroup = wsInput.Range(wsInput.Cells(row, 1), wsInput.Cells(row, MaxSrcCol))

            currentYear = y
            currentMonth = m
            currentDay = d
        Else
            ' Add the row To the current group
            Set currentGroup = Union(currentGroup, wsInput.Range(wsInput.Cells(row, 1), wsInput.Cells(row, MaxSrcCol)))
        End If
    Next row

    ' Process the last group of rows
    If Not currentGroup Is Nothing Then
        ProcessGroupByDay currentGroup, CStr(currentYear) & "/" & CStr(currentMonth) & "/" & CStr(currentDay), idxcolOutput

        ProcessGroupByDaySmallAmount currentGroup, CStr(currentYear) & "/" & CStr(currentMonth) & "/" & CStr(currentDay), idxcolSmallAmount

        ProcessGroupByDayDiff currentGroup, CStr(currentYear) & "/" & CStr(currentMonth) & "/" & CStr(currentDay), idxcolDiff
    End If

    Dim lastCol As Long
    Dim i       As Long
    Dim nDays   As Integer

    lastCol = wsInter.Cells(RowTask12SmallTransferXAxis, wsInter.columns.Count).End(xlToLeft).Column

    '  任務12, 一元交易
    nDays = 0
    For i = 1 To lastCol
        If wsInter.Cells(RowTask12SmallTransfer1Dollar, i).value >= 1 Then
            Utils.CountOne nDays
        End If
    Next i

    If nDays > 0 Then
        With Range(posTask12Result)
            .wrapText = True
            .value = RowTaskAlertPrefixByViewer & CStr(nDays) & RowTaskAlertPostfixBy1Dollar
        End With
    End If

    ExcelUtils.CreateBarChart RowTask12SmallTransferXAxis, RowTask12SmallTransfer1Dollar, _
                              Range(RowTask12TS1DollarRange), RowTask12TS1DollarTitle, RowTask12TS1DollarYLabel

    '  任務12, 百元交易
    nDays = 0
    For i = 1 To lastCol
        If wsInter.Cells(RowTask12SmallTransferHundred, i).value >= 1 Then
            Utils.CountOne nDays
        End If
    Next i

    If nDays > 0 Then
        With Range(posTask13Result)
            .wrapText = True
            .value = RowTaskAlertPrefixByViewer & CStr(nDays) & RowTaskAlertPostfixByHundred
        End With
    End If

    ExcelUtils.CreateBarChart RowTask12SmallTransferXAxis, RowTask12SmallTransferHundred, _
                              Range(RowTask12TSHundredRange), RowTask12TSHundredTitle, RowTask12TSHundredYLabel

    '  任務12, 跨行頻繁交易
    nDays = 0
    For i = 1 To lastCol
        If wsInter.Cells(RowTask12SmallTransferMany, i).value >= 1 Then
            Utils.CountOne nDays
        End If
    Next i

    If nDays > 0 Then
        With Range(posTask14Result)
            .wrapText = True
            .value = RowTaskAlertPrefixByViewer & CStr(nDays) & RowTaskAlertPostfixByMany
        End With
    End If

    ExcelUtils.CreateBarChart RowTask12SmallTransferXAxis, RowTask12SmallTransferMany, _
                              Range(RowTask12TSManyRange), RowTask12TSManyTitle, RowTask12TSManyYLabel

    ' 任務4, ??
    ' 當日現金提領_臨櫃
    ' 任務5, 可疑大額提領
    Dim nTran       As Long
    Dim strSusDates As String
    Dim hasSusDates As Boolean
    Dim dict        As New Scripting.Dictionary

    lastCol = wsInter.Cells(RowTask4WithdrawOverCounterXAxis, wsInter.columns.Count).End(xlToLeft).Column

    For i = 1 To lastCol
        Dim val, key As Variant
        key = wsInter.Cells(RowTask4WithdrawOverCounterXAxis, i).value
        val = wsInter.Cells(RowTask4WithdrawOverCounterYAxis, i).value

        If val > 1 Then
            strSusDates = strSusDates & key & ": " & val & vbNewLine
            hasSusDates = True
        End If

        If Not dict.Exists(val) Then
            dict.Add val, 1
        Else
            dict(val) = dict(val) + 1
        End If
    Next i

    Dim col      As Long
    Dim cntTotal As Long

    col = 1
    For i = 1 To lastCol
        If dict.Exists(i) Then
            wsInter.Cells(RowTask5WithdrawAmountXAxis, col).value = RowTaskGeneralPrefix & CStr(i) & RowTaskGeneralPostfix
            wsInter.Cells(RowTask5WithdrawAmountYAxis, col).value = dict(i)
            Utils.CountOne col
            cntTotal = cntTotal + dict(i)
        End If
    Next i

    ExcelUtils.CreateBarChart RowTask5WithdrawAmountXAxis, RowTask5WithdrawAmountYAxis, _
                              Range(RowTask5WithdrawAmountRange), RowTask5WithdrawAmountTitle, RowTask5WithdrawAmountYLabel

    ' 任務 4, 可疑提領日
    With Range(posTask4Result)
        .value = ""

        If hasSusDates = True Then
            .wrapText = True
            .value = RowTaskAlertPrefixByWDrawer & vbNewLine & strSusDates
        End If
    End With

    ' 任務5, 可疑多間分行提領
    ' Specify the row index you want To traverse
    Dim idxCol    As Integer
    Dim hasMulti  As Boolean
    Dim strHits   As String
    Dim cellKey   As String
    Dim cellValue As String

    ' Loop through each cell in the specified row
    For idxCol = 1 To wsInter.Cells(RowTask5WithdrawBranchXAxis, wsInter.columns.Count).End(xlToLeft).Column
        ' Get the value of the current cell
        cellKey = wsInter.Cells(RowTask5WithdrawBranchXAxis, idxCol).value
        cellValue = Trim(wsInter.Cells(RowTask5WithdrawBranchYAxis, idxCol).value)
        If cellValue <> "" Then
            ' Process the cell value (replace With your code)
            ' Debug.Print "Column: " & idxCol & ", Value: " & cellValue
            hasMulti = True
            strHits = strHits & cellKey & ": " & cellValue & vbNewLine
        End If
    Next idxCol

    With Range(posTask5Result)
        .value = ""

        If hasMulti = True Then
            .wrapText = True
            .value = RowTaskAlertPrefixByWDrawer & vbNewLine & strHits
        End If
    End With

    ' 任務7, "金額差距甚小，疑為人頭過水帳戶"
    lastCol = wsInter.Cells(RowTask7DiffXAxis, wsInter.columns.Count).End(xlToLeft).Column

    Dim cntDiffLess   As Double
    Dim totalDeposit  As Double
    Dim totalTransfer As Double
    Dim ratio         As Double

    For i = 1 To lastCol
        totalDeposit = totalDeposit + CDbl(wsInter.Cells(RowTask7DiffDeposit, i).value)
        totalTransfer = totalTransfer + CDbl(wsInter.Cells(RowTask7DiffTransfer, i).value)

        ratio = CDbl(wsInter.Cells(RowTask7DiffRatio, i).value)
        If IsRatioDiffLess(ratio) = True Then
            Utils.CountOne cntDiffLess
        End If
    Next i

    If totalDeposit <> 0 Then
        ratio = Abs((totalDeposit - totalDeposit) / totalDeposit)
    Else
        ratio = 0
    End If

    With Range(posTask7_1Result)
        .value = ""

        If IsRatioDiffLess(ratio) = True Then
            .value = RowTaskAlertPrefixByViewer & CStr(ratio) & RowTaskAlertPostfixByRatio
        End If
    End With

    With Range(posTask7_2Result)
        .value = ""

        If cntDiffLess > 0 Then
            .value = RowTaskAlertPrefixByViewer & CStr(cntDiffLess) & RowTaskAlertPostfixByDiffRatio
        End If
    End With

    ExcelUtils.CreateBarChart RowTask7DiffXAxis, RowTask7DiffRatio, _
                              Range(RowTask7DiffRatioRange), RowTask7DiffRatioTitle, RowTask7DiffRatioYLabel
End Sub

Function ProcessGroupByDay(groupRange As Range, colName As String, idxOut As Long) As Boolean
    Dim intResult     As Long
    Dim sum           As Long
    Dim cell          As Range
    Dim wsInter       As Worksheet
    Dim cntSuspicious As Long
    Dim cntBranch     As Long
    Dim dictBranch    As New Scripting.Dictionary
    Dim strBranch     As String
    Dim strKey        As String
    Dim result        As Variant

    Set wsInter = ThisWorkbook.Worksheets(SheetNameIntermediate)
    intResult = groupRange.row + groupRange.Rows.Count - 1

    ' 任務 4 & 任務 5
    ' 當日現金提領_臨櫃
    For Each cell In groupRange.columns(ColShInDataBranchID).Cells
        If ExcelUtils.IsAutoTSID(cell.Text) = True Then
            result = LookupBranchInfo(cell.Text)
            If IsError(result) Then
                ' Debug.Print "branch info Not found" & cell.Text
                strKey = CStr(cell.Text)
            Else
                strKey = CStr(result)
            End If

            If Not dictBranch.Exists(strKey) Then
                Utils.CountOne cntBranch
                dictBranch.Add strKey, 1
                strBranch = strBranch + strKey + ", "
            End If
        End If
    Next cell

    If Len(strBranch) > 1 Then
        strBranch = Left(strBranch, Len(strBranch) - 2)
    End If

    ' 任務 4
    ' 當日現金提領_臨櫃
    ' 計算交易金額刻意維持在 50 萬以下接近 50 萬元間 (45 萬 ~ 50 萬)
    Dim cntSuspcious As Long ' 可疑次數
    For Each cell In groupRange.columns(ColShInDataAmtDraw).Cells
        If IsNumeric(cell.value) Then
            If (cell.value >= RowTask4WithdrawAmountLBSusAmt) And _
               (cell.value < RowTask4WithdrawAmountUBSusAmt) Then
                Utils.CountOne cntSuspicious
                ExcelUtils.HighlightRow SheetNameSimple, cell.row, ReasonCloseTo50w, ColorRed
            End If
        End If
    Next cell

    If cntSuspicious > 0 Then
        wsInter.Cells(RowTask4WithdrawOverCounterXAxis, idxOut).value = colName
        wsInter.Cells(RowTask4WithdrawOverCounterYAxis, idxOut).value = cntSuspicious

        wsInter.Cells(RowTask5WithdrawBranchXAxis, idxOut).value = colName
        If cntBranch >= RowTask5WithdrawBranchLBCnt Then
            wsInter.Cells(RowTask5WithdrawBranchYAxis, idxOut).value = strBranch
        End If

        ProcessGroupByDay = True
        Exit Function
    End If

    ProcessGroupByDay = False
End Function

Function ProcessGroupByDaySmallAmount(groupRange As Range, colName As String, idxOut As Long) As Boolean
    Dim intResult As Long
    Dim cell      As Range
    Dim amount    As Long
    Dim wsInter   As Worksheet

    Dim cntOneDollar, cntHundred, cntInterBank As Long

    Set wsInter = ThisWorkbook.Worksheets(SheetNameIntermediate)
    intResult = groupRange.row + groupRange.Rows.Count - 1

    ' 任務 12 & 13
    ' 一元交易 & 小於 100 元交易
    For Each cell In groupRange.columns(ColShInDataTSType).Cells
        amount = groupRange.Cells(cell.row - groupRange.row + 1, ColShInDataAmtDraw).value

        'FIXME: 跨行存款怎麼看
        If (cell.Text = ColTSSummaryVal03) And (amount > 0) Then
            Utils.CountOne cntInterBank
        End If

        ' 跨行轉帳 or 本行轉帳
        If cell.Text = ColTSSummaryVal03 Or cell.Text = ColTSSummaryVal04 Then
            ' 視為正常轉帳
            If amount >= 100 Then
                ' 大於或等於 100 元的轉帳
                ' 不列入計算
            End If

            ' 小於 100 元交易
            If (amount < 100) And (amount > 1) Then
                Utils.CountOne cntHundred
                ExcelUtils.HighlightRow SheetNameSimple, cell.row, ReasonSmallTSAmount, ColorPink
            End If

            ' 1 元交易
            If amount = 1 Then
                Utils.CountOne cntOneDollar
                ExcelUtils.HighlightRow SheetNameSimple, cell.row, ReasonSmallTSAmount, ColorRed
            End If

        End If
    Next cell

    If cntOneDollar > 0 Then
        wsInter.Cells(RowTask12SmallTransferXAxis, idxOut).value = colName
        wsInter.Cells(RowTask12SmallTransfer1Dollar, idxOut).value = cntOneDollar
    End If

    If cntHundred > 0 Then
        wsInter.Cells(RowTask12SmallTransferXAxis, idxOut).value = colName
        wsInter.Cells(RowTask12SmallTransferHundred, idxOut).value = cntHundred
    End If

    If cntInterBank >= RowTask12SmallTSLBMany Then ' 跨行交易分行數閥值
        wsInter.Cells(RowTask12SmallTransferXAxis, idxOut).value = colName
        wsInter.Cells(RowTask12SmallTransferMany, idxOut).value = cntInterBank

    End If

    If cntOneDollar > 0 Or cntHundred > 0 Or cntInterBank >= RowTask12SmallTSLBMany Then
        ProcessGroupByDaySmallAmount = True
        Exit Function
    End If

    ProcessGroupByDaySmallAmount = False
End Function

Function ProcessGroupByDayDiff(groupRange As Range, colName As String, idxOut As Long) As Boolean
    Dim cell        As Range
    Dim wsInter     As Worksheet
    Dim sumDeposit  As Double
    Dim sumTransfer As Double
    Dim amtDeposit  As Double
    Dim amtTransfer As Double

    Set wsInter = ThisWorkbook.Worksheets(SheetNameIntermediate)

    ' 任務 12 & 任務 13
    ' 一元交易 & 小於 100 元交易
    For Each cell In groupRange.columns(ColShInDataAmtDraw).Cells
        amtTransfer = CDbl(cell.value)
        amtDeposit = CDbl(groupRange.Cells(cell.row - groupRange.row + 1, ColShInDataAmtDeposit).value)

        sumTransfer = sumTransfer + amtTransfer
        sumDeposit = sumDeposit + amtDeposit
    Next cell

    'FIXME: set format as percentage!
    wsInter.Cells(RowTask7DiffXAxis, idxOut).value = colName
    wsInter.Cells(RowTask7DiffDeposit, idxOut).value = sumDeposit
    wsInter.Cells(RowTask7DiffTransfer, idxOut).value = sumTransfer
    With wsInter.Cells(RowTask7DiffRatio, idxOut)
        .value = 0

        If sumDeposit <> 0 Then
            .value = Abs((sumDeposit - sumTransfer) / sumDeposit)
        End If
    End With

    ProcessGroupByDayDiff = True
End Function

Sub GetVisitedBranch()
    Dim wsInput    As Worksheet
    Dim wsInter    As Worksheet
    Dim dictBranch As New Scripting.Dictionary
    Dim lastRow    As Long
    Dim row        As Long
    Dim cell       As Range

    Set wsInput = ThisWorkbook.Worksheets(SheetNameInputData)
    Set wsInter = ThisWorkbook.Worksheets(SheetNameIntermediate)
    lastRow = ExcelUtils.GetLastRowNumber(wsInput)

    ' Loop through rows starting from row rowDataBegin, day by day
    For row = RowDataBegin To lastRow        ' lastRow FIXME: For debug only
        Set cell = wsInput.Cells(row, ColShInDataBranchID)

        If ExcelUtils.IsAutoTSID(cell.Text) = False Then
            If Not dictBranch.Exists(cell.Text) Then
                dictBranch.Add cell.Text, 1
            Else
                dictBranch(cell.Text) = dictBranch(cell.Text) + 1
            End If
        End If
    Next row

    ' Convert the dictionary To arrays
    Dim result        As Variant
    Dim strKey        As String
    Dim keysArray()   As Variant
    Dim valuesArray() As Variant
    Dim key           As Variant

    ReDim keysArray(0)
    ReDim valuesArray(0)
    If dictBranch.Count > 0 Then
        ReDim keysArray(1 To dictBranch.Count)
        ReDim valuesArray(1 To dictBranch.Count)

        Dim i As Long
        i = 1
        For Each key In dictBranch.keys
            keysArray(i) = key
            valuesArray(i) = dictBranch(key)
            Utils.CountOne i
        Next key

        ' Sort the arrays by values
        Utils.QuickSort valuesArray, keysArray, LBound(valuesArray), UBound(valuesArray)

        ' Write the sorted arrays
        For i = LBound(valuesArray) To UBound(valuesArray)
            result = LookupBranchInfo(keysArray(i))
            If IsError(result) Then
                Debug.Print "real branch info Not found" & cell.Text
                'strKey = CStr(keysArray(i))
                strKey = CStr(Format$(keysArray(i), "0000"))
            Else
                strKey = CStr(result)
            End If

            wsInter.Cells(RowTask5VisitedBranchXAxis, i).value = RowTask5VBCodePrefix & Format(strKey, "0000")
            wsInter.Cells(RowTask5VisitedBranchYAxis, i).value = valuesArray(i)
        Next i
    End If

    ExcelUtils.CreateBarChart RowTask5VisitedBranchXAxis, RowTask5VisitedBranchYAxis, _
                              Range(RowTask5VisitedBranchRange), RowTask5VisitedBranchTitle, RowTask5VisitedBranchYLabel
End Sub

Sub GetTransferAmount()
    Dim wsInput           As Worksheet
    Dim wsInter           As Worksheet
    Dim dictTransfer      As New Scripting.Dictionary
    Dim dictTransferCount As New Scripting.Dictionary
    Dim dictDeposit       As New Scripting.Dictionary
    Dim lastRow           As Long
    Dim row               As Long
    Dim cell              As Range
    Dim val               As Double

    Set wsInput = ThisWorkbook.Worksheets(SheetNameInputData)
    Set wsInter = ThisWorkbook.Worksheets(SheetNameIntermediate)
    lastRow = ExcelUtils.GetLastRowNumber(wsInput)

    ExcelUtils.ApplyFontStyle wsInter

    ' Loop through rows starting from row rowDataBegin, day by day
    For row = RowDataBegin To lastRow
        Set cell = wsInput.Cells(row, ColShInDataAccount)

        If Not IsEmpty(cell) Then
            val = CDbl(wsInput.Cells(row, ColShInDataAmtDraw).value)
            If val > 0 Then
                If Not dictTransfer.Exists(cell.Text) Then
                    dictTransfer.Add cell.Text, val
                    dictTransferCount.Add cell.Text, 1
                Else
                    dictTransfer(cell.Text) = dictTransfer(cell.Text) + val
                    dictTransferCount(cell.Text) = dictTransferCount(cell.Text) + 1
                End If
            End If

            val = CDbl(wsInput.Cells(row, ColShInDataAmtDeposit).value)
            If val > 0 Then
                If Not dictDeposit.Exists(cell.Text) Then
                    dictDeposit.Add cell.Text, val
                Else
                    dictDeposit(cell.Text) = dictDeposit(cell.Text) + 1
                End If
            End If
        End If
    Next row

    If dictTransfer.Count <= 0 Then
        GoTo EndOfSub
    End If

    ' Convert the dictionary To arrays
    Dim result        As Variant
    Dim strKey        As String
    Dim keysArray()   As Variant
    Dim valuesArray() As Variant
    Dim key           As Variant
    Dim i             As Long

    ReDim keysArray(1 To dictTransfer.Count)
    ReDim valuesArray(1 To dictTransfer.Count)

    i = 1
    For Each key In dictTransfer.keys
        keysArray(i) = key
        valuesArray(i) = dictTransfer(key)
        Utils.CountOne i
    Next key

    ' Sort the arrays by values
    Utils.QuickSort valuesArray, keysArray, LBound(valuesArray), UBound(valuesArray)

    ' Write the sorted arrays
    For i = LBound(valuesArray) To UBound(valuesArray)
        ExcelUtils.ApplyCellsValue wsInter, keysArray(i), ForceStringFormat, xlLeft, (RowTopCounterparty + i - 1), 1
        ExcelUtils.ApplyCellsValue wsInter, valuesArray(i), NumberFormat, xlRight, (RowTopCounterparty + i - 1), 2
        ExcelUtils.ApplyCellsValue wsInter, dictTransferCount(keysArray(i)), NumberFormat, xlRight, (RowTopCounterparty + i - 1), 3
    Next i

    ReDim keysArray(1 To dictDeposit.Count)
    ReDim valuesArray(1 To dictDeposit.Count)

    i = 1
    For Each key In dictDeposit.keys
        keysArray(i) = key
        valuesArray(i) = dictDeposit(key)
        Utils.CountOne i
    Next key

    ' Sort the arrays by values
    Utils.QuickSort valuesArray, keysArray, LBound(valuesArray), UBound(valuesArray)

    ' Write the sorted arrays
    For i = LBound(valuesArray) To UBound(valuesArray)
        ExcelUtils.ApplyCellsValue wsInter, keysArray(i), ForceStringFormat, xlLeft, (RowTopCounterparty + i - 1), 4
        ExcelUtils.ApplyCellsValue wsInter, valuesArray(i), NumberFormat, xlRight, (RowTopCounterparty + i - 1), 5
    Next i
EndOfSub:
    ' FIXME: 區分跨行
End Sub

'/* 如何判斷是ATM交易:  1. 看櫃員代碼 (慢) 2. 收款機/ATM 欄 */
' or cash
' TASK ATM 存現 地緣性 by day
Sub ProcessRowsATM()
    Dim wsInput      As Worksheet
    Dim wsInter      As Worksheet
    Dim lastRow      As Long
    Dim lastCol      As Long
    Dim currentGroup As Range
    Dim row          As Long
    Dim idxcolOutput As Integer
    Dim idxcolQuick  As Integer

    Dim currentYear, currentMonth, currentDay As Integer

    idxcolOutput = 1  ' output column index
    idxcolQuick = 1

    Set wsInput = ThisWorkbook.Worksheets(SheetNameInputData)
    Set wsInter = ThisWorkbook.Worksheets(SheetNameIntermediate)
    lastRow = ExcelUtils.GetLastRowNumber(wsInput)
    lastCol = wsInter.Cells(RowTask9ATMDepositXAxis, wsInter.columns.Count).End(xlToLeft).Column

    ' Loop through rows starting from row rowDataBegin, day by day
    For row = RowDataBegin To lastRow
        Dim dateValue As Date
        Dim y, m, d   As Integer
        Dim ret       As Boolean

        dateValue = wsInput.Cells(row, ColShInDataTSDate).value
        y = year(dateValue)
        m = month(dateValue)
        d = Day(dateValue)
        ' Check If the date has changed Or If it's the first row
        If (row = RowDataBegin) Or (y <> currentYear) Or (m <> currentMonth) Or (d <> currentDay) Then
            ' Process the current group of rows
            If Not currentGroup Is Nothing Then
                ret = ProcessGroupByATM(currentGroup, CStr(currentYear) & "/" & CStr(currentMonth) & "/" & CStr(currentDay), idxcolOutput)
                If ret = True Then
                    idxcolOutput = idxcolOutput + 1
                End If
                ret = ProcessGroupByCashQuickInOut(currentGroup, CStr(currentYear) & "/" & CStr(currentMonth) & "/" & CStr(currentDay), idxcolQuick)
                If ret = True Then
                    idxcolQuick = idxcolQuick + 1
                End If

                Set currentGroup = Nothing
            End If

            ' Initialize a New group For the current date
            Set currentGroup = wsInput.Range(wsInput.Cells(row, 1), wsInput.Cells(row, MaxSrcCol))

            currentYear = y
            currentMonth = m
            currentDay = d
        Else
            ' Add the row To the current group
            Set currentGroup = Union(currentGroup, wsInput.Range(wsInput.Cells(row, 1), wsInput.Cells(row, MaxSrcCol)))
        End If
    Next row

    ' Process the last group of rows
    If Not currentGroup Is Nothing Then
        ProcessGroupByATM currentGroup, CStr(currentYear) & "/" & _
                          CStr(currentMonth) & "/" & CStr(currentDay), idxcolOutput
        ProcessGroupByCashQuickInOut currentGroup, CStr(currentYear) & "/" & _
                          CStr(currentMonth) & "/" & CStr(currentDay), idxcolQuick
    End If

    ' 任務 9, 任務 10
    ' 當日現金提領/存入_ATM
    ' "同時段多點小額存入、一地提領，疑為車手態樣 >= 3 ATM" & "一天內多點存現但不具地緣性 >= 10筆 "
    ' RowTask9ATMDepositLBLoc & RowTask9ATMDepositLBCnt
    cntMultiATM = 0
    cntMultiDeposit = 0

    Dim i As Long
    For i = 1 To lastCol
        If wsInter.Cells(RowTask9ATMDepositYAxis, i) >= RowTask9ATMDepositLBCnt Then
            Utils.CountOne cntMultiDeposit
        End If

        If wsInter.Cells(RowTask9ATMDepositLoc, i) >= RowTask9ATMDepositLBLoc Then
            Utils.CountOne cntMultiATM
        End If
    Next i

    With Range(posTask9Result)
        .value = ""

        If cntMultiATM > 0 Then
            .wrapText = True
            .value = RowTaskAlertPrefixByViewer & CStr(cntMultiATM) & RowTaskGeneralPostfix & _
                     RowTaskGeneralSingleDay & RowTask9ATMDepositLBCnt & RowTaskAlertPostfixByATMManyLoc
        End If
    End With

    With Range(posTask10Result)
        .value = ""

        If cntMultiDeposit > 0 Then
            .wrapText = True
            .value = RowTaskAlertPrefixByViewer & CStr(cntMultiDeposit) & RowTaskGeneralPostfix & _
                     RowTaskGeneralSingleDay & RowTask9ATMDepositLBLoc & RowTaskAlertPostfixByATMDeposit
        End If
    End With

    ExcelUtils.CreateBarChart RowTask9ATMDepositXAxis, RowTask9ATMDepositYAxis, _
                              Range(RowTask9ATMDepositRange), RowTask9ATMDepositTitle, RowTask9ATMDepositYLabel

    lastCol = wsInter.Cells(RowTask6QuickXAxis, wsInter.columns.Count).End(xlToLeft).Column

    With Range(posTask6Result)
        .value = ""

        If lastCol > 0 Then
            .wrapText = True
            .value = RowTaskAlertPrefixByViewer & CStr(lastCol) & RowTaskAlertPostfixByQuickInOut
        End If
    End With

    ExcelUtils.CreateBarChart RowTask6QuickXAxis, RowTask6QuickYAxis, _
                              Range(RowTask6QuickRange), RowTask6QuickTitle, RowTask6QuickYLabel
End Sub

Function ProcessGroupByATM(groupRange As Range, colName As String, idxOut As Integer) As Boolean
    Dim cell        As Range
    Dim wsInter     As Worksheet
    Dim dictATM     As New Scripting.Dictionary
    Dim cntATM      As Integer
    Dim strChannel  As String
    Dim strBankID   As String
    Dim strBranch   As String
    Dim cntATMTrans As Long
    Dim amount      As Double

    Set wsInter = ThisWorkbook.Worksheets(SheetNameIntermediate)

    For Each cell In groupRange.columns(ColShInDataBranchID).Cells
        If ExcelUtils.IsAutoTSID(cell.Text) = True Then
            strChannel = groupRange.Cells(cell.row - groupRange.row + 1, ColShInDataChannel).Text
            amount = CDbl(groupRange.Cells(cell.row - groupRange.row + 1, ColShInDataAmtDeposit).value)

            ' Check deposit amount > 0 only
            If amount > 0 Then
                'If strChannel = "存款機" Or strChannel = "ＡＴＭ" Or strChannel = "ATM" Then
                If Utils.ArrayContainsItem(ExcelUtils.GetATMChannelArray(), strChannel) = True Then
                    Utils.CountOne GotATM
                    Utils.CountOne cntATMTrans

                    If Not dictATM.Exists(strChannel) Then
                        dictATM.Add strChannel, 1
                    Else
                        dictATM(strChannel) = dictATM(strChannel) + 1
                    End If
                ElseIf Len(strChannel) < 4 And Utils.ContainsNumbers(strChannel) = True Then
                    Utils.CountOne GotATM
                    Utils.CountOne cntATMTrans

                    strBankID = Format(CInt(strChannel), "000")

                    If Not dictATM.Exists(strBankID) Then
                        dictATM.Add strBankID, 1
                    Else
                        dictATM(strBankID) = dictATM(strBankID) + 1
                    End If
                End If
            End If
        End If
ContinueLoop:
    Next cell

    Dim result As String
    If (dictATM.Count >= RowTask9ATMDepositLBLoc) Or (cntATMTrans >= RowTask9ATMDepositLBCnt) Then
        result = Join(dictATM.keys, ";")
        wsInter.Cells(RowTask9ATMDepositXAxis, idxOut).value = colName
        wsInter.Cells(RowTask9ATMDepositYAxis, idxOut).value = cntATMTrans
        wsInter.Cells(RowTask9ATMDepositLoc, idxOut).value = dictATM.Count
        wsInter.Cells(RowTask9ATMDepositLocDetail, idxOut).value = result
        'Debug.Print "Write " & colName & " " & CStr(cntATMTrans) & " " & CStr(dictATM.Count) & " " & result

        ProcessGroupByATM = True
        Exit Function
    End If

    ProcessGroupByATM = False
End Function

Function ProcessGroupByCashQuickInOut(groupRange As Range, colName As String, idxOut As Integer) As Boolean
    Dim cell        As Range
    Dim wsInter     As Worksheet
    Dim currentTime As Date
    Dim bInit       As Boolean
    Dim t           As Date
    Dim tDiff       As Double
    Dim bQuick      As Boolean
    Dim cntQuick    As Integer
    Dim amtTransfer As Double
    Dim amtDeposit  As Double
    Dim bTransfer   As Boolean
    Dim bDeposit    As Boolean

    Set wsInter = ThisWorkbook.Worksheets(SheetNameIntermediate)

    For Each cell In groupRange.columns(ColShInDataTSType).Cells
        If cell.Text = ColTSSummaryVal01 Then
            amtTransfer = CDbl(groupRange.Cells(cell.row - groupRange.row + 1, ColShInDataAmtDraw).value)
            If amtTransfer > 0 Then
                bTransfer = True
            End If

            amtDeposit = CDbl(groupRange.Cells(cell.row - groupRange.row + 1, ColShInDataAmtDeposit).value)
            If amtDeposit > 0 Then
                bDeposit = True
            End If

            If bInit = True Then
                t = groupRange.Cells(cell.row - groupRange.row + 1, ColShInDataTSTime).value
                tDiff = DateDiff("n", currentTime, t)
                If t <= RowTask6QuickLBMins Then
                    ExcelUtils.HighlightRow SheetNameSimple, cell.row, ReasonTSFast, ColorRed
                    If bQuick = False Then
                        bQuick = True
                        ExcelUtils.HighlightRow SheetNameSimple, (cell.row - 1), ReasonTSFast, ColorRed
                        cntQuick = cntQuick + 2
                    Else
                        cntQuick = cntQuick + 1
                    End If
                Else
                    currentTime = t
                    bQuick = False
                End If
            Else
                currentTime = groupRange.Cells(cell.row - groupRange.row + 1, ColShInDataTSTime).value
                bInit = True
            End If
        End If
    Next cell

    If (cntQuick > 1) And (bTransfer = True) And (bDeposit = True) Then
        wsInter.Cells(RowTask6QuickXAxis, idxOut).value = colName
        wsInter.Cells(RowTask6QuickYAxis, idxOut).value = cntQuick
        ProcessGroupByCashQuickInOut = True
        Exit Function
    End If

    ProcessGroupByCashQuickInOut = False
End Function

Sub ProcessRowsBalance()
    Dim wsInput        As Worksheet
    Dim wsInter        As Worksheet
    Dim lastRow        As Long
    Dim currentDate    As Date
    Dim currentBalance As Long
    Dim targetCol      As Long
    Dim i              As Long

    Set wsInput = ThisWorkbook.Sheets(SheetNameInputData)
    Set wsInter = ThisWorkbook.Worksheets(SheetNameIntermediate)

    lastRow = ExcelUtils.GetLastRowNumber(wsInput)

    ' Initialize variables
    currentBalance = 0
    currentDate = wsInput.Cells(RowDataBegin, ColShInDataTSDate).value        ' Assuming the first date is in A9
    If VarType(wsInput.Cells(RowDataBegin, ColShInDataBalance).value) <> vbString Then
        currentBalance = wsInput.Cells(RowDataBegin, ColShInDataBalance).value        ' Assuming the first balance is in I9
    End If

    targetCol = 1
    ' Loop through each row starting from the second row
    For i = RowDataBegin To lastRow
        If VarType(wsInput.Cells(i, ColShInDataBalance).value) <> vbString Then        ' workaround for data error (欄位飄掉)
            ' Check if the date is the same
            If wsInput.Cells(i, ColShInDataTSDate).value = currentDate Then
                ' Update the current balance
                currentBalance = CDbl(wsInput.Cells(i, ColShInDataBalance).value)
            Else
                ' Write the date and balance to the new rows
                wsInter.Cells(RowTask11BalanceXAxis, targetCol).value = currentDate
                wsInter.Cells(RowTask11BalanceYAxis, targetCol).value = currentBalance

                ' Update currentDate and currentBalance for the new date
                currentDate = wsInput.Cells(i, ColShInDataTSDate).value
                currentBalance = CDbl(wsInput.Cells(i, ColShInDataBalance).value)
                ' Move to the next row
                targetCol = targetCol + 1
            End If
        End If
    Next i
    ' Write the last date and balance to the final row

    wsInter.Cells(RowTask11BalanceXAxis, targetCol).value = currentDate
    wsInter.Cells(RowTask11BalanceYAxis, targetCol).value = currentBalance

    ExcelUtils.CreateBarChart RowTask11BalanceXAxis, RowTask11BalanceYAxis, _
                              Range(RowTask11BalanceRange), RowTask11BalanceTitle, RowTask11BalanceYLabel

    ' Compute the date difference
    Dim startDate         As Date
    Dim endDate           As Date
    Dim dayDiff, dayTotal As Long
    Dim idxDate           As Integer
    Dim currentValue      As Long
    Dim blUnderThreshold  As Boolean
    Dim cntUnderDays      As Integer
    Dim lastCol           As Integer

    lastCol = wsInter.Cells(RowTask11BalanceXAxis, wsInput.columns.Count).End(xlToLeft).Column

    startDate = wsInter.Cells(RowTask11BalanceXAxis, 1).value
    endDate = wsInter.Cells(RowTask11BalanceXAxis, lastCol).value
    dayDiff = DateDiff("d", startDate, endDate)

    dayTotal = dayDiff + 1
    idxDate = 1

    cntUnderDays = 0
    currentDate = wsInter.Cells(RowTask11BalanceXAxis, idxDate).value
    currentValue = wsInter.Cells(RowTask11BalanceYAxis, idxDate).value
    If currentValue <= RowTask11BalanceUBCnt Then
        blUnderThreshold = True
    End If

    Dim d As Date
    For i = 1 To dayTotal
        d = DateAdd("d", (i - 1), startDate)
        If blUnderThreshold = True Then
            cntUnderDays = cntUnderDays + 1
            Utils.CountOne cntUnderDays
        End If

        If d = currentDate Then
            ' matched
            'idxDate = idxDate + 1
            Utils.CountOne idxDate
            currentDate = wsInter.Cells(RowTask11BalanceXAxis, idxDate).value
            currentValue = wsInter.Cells(RowTask11BalanceYAxis, idxDate).value
            If currentValue <= RowTask11BalanceUBCnt Then
                blUnderThreshold = True
            Else
                blUnderThreshold = False
            End If
        End If
    Next i

    ' 任務 11, 當日存款餘額
    With Range(posTask11Result)
        .value = ""

        If cntUnderDays > 0 Then
            .wrapText = True
            .value = RowTaskAlertPrefixByViewer & CStr(dayTotal) & RowTaskGeneralDay & _
                     RowTaskGeneralPrefix & CStr(cntUnderDays) & RowTaskAlertPostfixByAmtLess1K
            ' Debug.Print "total days Is " & CStr(dayTotal) & " under days Is " & CStr(cntUnderDays)
        End If
    End With
End Sub

Private Sub PreparingData()
    Application.ErrorCheckingOptions.NumberAsText = False
    Application.ScreenUpdating = False ' For performance
    Application.DisplayAlerts = False

    ' 準備虛擬帳號跟警示戶資料
    ExtraData.InitDicts

    ' 準備資料, 正規化 1原始資料 成為 2.2清整後資料
    NormalizeRawData

    ' 分析資料
    ' 資料分析處理後，處理結果放置於暫存區
    Analyze

    ' 讀出處理結果，準備3.2 金流及交易對手 頁面
    CreatePivotTables

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True ' For performance
    Application.ErrorCheckingOptions.NumberAsText = True
End Sub

Public Sub BtnAnalyze_Click()
    Dim wsMain  As Worksheet
    Dim wsMoney As Worksheet
    Set wsMain = ThisWorkbook.Sheets(SheetNameMain)
    Set wsMoney = ThisWorkbook.Sheets(SheetNameMoney)
    With wsMain.Range(posRunStatus)
        .NumberFormat = ForceStringFormat
        .Font.name = FontName
        .Font.size = AnalyBtnFontSize
        .value = BtnStatusStartRun
    End With

    PreparingData

    With wsMain.Range(posRunStatus)
        .NumberFormat = ForceStringFormat
        .Font.name = FontName
        .Font.size = AnalyBtnFontSize
        .value = BtnStatusStopRun
    End With

    ExcelUtils.ActiveSheet wsMoney
End Sub

Private Sub Analyze()
    Dim wsMain   As Worksheet
    Dim wsSimple As Worksheet
    Dim wsInput  As Worksheet

    ' 2.1主頁面
    Set wsMain = ThisWorkbook.Sheets(SheetNameMain)
    ' 3.1交易明細
    Set wsSimple = ThisWorkbook.Sheets(SheetNameSimple)
    ' 2.2清整後資料
    Set wsInput = ThisWorkbook.Sheets(SheetNameInputData)

    ' 清掉所有資料
    ExcelUtils.RemoveAllCharts SheetNameMain
    wsMain.Range(RowShMainContentRange).ClearContents

    ExcelUtils.ClearSheet SheetNameIntermediate
    ExcelUtils.ClearSheet SheetNameSimple, False
    ExcelUtils.ClearSheet SheetNameMoney

    ProcessRowsByMonth
    ProcessRowsByDay
    GetVisitedBranch
    ProcessRowsATM
    ProcessRowsBalance

    ' Assume source data is ready
    ' 從2.2清整後資料裡，轉換資料到3.1交易明細
    ' 以下適用得到的欄位
    FillSimplePage ExcelUtils.GetShInDataColsToShSimple()

    GetTransferAmount

    ExcelUtils.ApplyFontStyle wsMain, False, CharAreaFontSize, False, "A", "A"
    ExcelUtils.ApplyFontStyle wsMain, False, CharAreaFontSize, True, "B", "E"
End Sub

' prepare data in the simple page, put all UI related here
' 這個子函式用來做出 3.1交易明細
Sub FillSimplePage(arrColumnIDs() As Variant)
    Dim wsSrc   As Worksheet
    Dim wsDest  As Worksheet
    Dim colID   As Variant
    Dim lastRow As Long
    Dim destCol As Long
    destCol = 1

    ' Set the source and destination worksheets
    Set wsSrc = ThisWorkbook.Sheets(SheetNameInputData)
    Set wsDest = ThisWorkbook.Sheets(SheetNameSimple)

    ' Start duplicating columns from the 7th column in the destination sheet
    lastRow = ExcelUtils.GetLastRowNumber(wsSrc)

    ExcelUtils.ApplySheetSimpleStyle

    ' Copy source columns
    For Each colID In arrColumnIDs
    'arrDup = Array(1, 16, 4, 17, 18, 9, 10, 14, 13, 27)
        If colID >= 1 And colID <= wsSrc.columns.Count Then
            wsDest.Range( _
                wsDest.Cells(RowHeaderBegin, destCol), _
                wsDest.Cells(lastRow, destCol)).value = _
                wsSrc.Range(wsSrc.Cells(RowHeaderBegin, colID), wsSrc.Cells(lastRow, colID)).value

            wsDest.columns(destCol).AutoFit
            Utils.CountOne destCol
        End If
    Next colID

    ' used in checking virtual accounts and channels
    ' 暫時性資料， Y 欄紀錄交易通路；Z 欄紀錄銀行代碼
    wsDest.Range(wsDest.Cells(RowHeaderBegin, ColShSimpleTempTSCh), wsDest.Cells(lastRow, ColShSimpleTempTSCh)).value = _
                wsSrc.Range(wsSrc.Cells(RowHeaderBegin, ColShInDataTSChannel), wsSrc.Cells(lastRow, ColShInDataTSChannel)).value
    wsDest.Range(wsDest.Cells(RowHeaderBegin, ColShSimpleTempCode), wsDest.Cells(lastRow, ColShSimpleTempCode)).value = _
                wsSrc.Range(wsSrc.Cells(RowHeaderBegin, ColShInDataBankCode), wsSrc.Cells(lastRow, ColShInDataBankCode)).value

    ' 暫時性資料， AA 欄紀錄照會帳戶, 虛擬帳號對應公司名；AB 欄紀錄標誌虛擬原因
    '             AC 欄紀錄警示帳戶
    wsDest.Range(wsDest.Cells(RowHeaderBegin, ColShSimpleTempVAccName), wsDest.Cells(lastRow, ColShSimpleTempVAccName)).value = _
                wsSrc.Range(wsSrc.Cells(RowHeaderBegin, ColShInDataVAccCName), wsSrc.Cells(lastRow, ColShInDataVAccCName)).value
    wsDest.Range(wsDest.Cells(RowHeaderBegin, ColShSimpleTempVAccReason), wsDest.Cells(lastRow, ColShSimpleTempVAccReason)).value = _
                wsSrc.Range(wsSrc.Cells(RowHeaderBegin, ColShInDataVAccReason), wsSrc.Cells(lastRow, ColShInDataVAccReason)).value
    wsDest.Range(wsDest.Cells(RowHeaderBegin, ColShSimpleTempWAccName), wsDest.Cells(lastRow, ColShSimpleTempWAccName)).value = _
                wsSrc.Range(wsSrc.Cells(RowHeaderBegin, ColShInDataWAccCName), wsSrc.Cells(lastRow, ColShInDataWAccCName)).value

    ' 設定其他3.1交易明細特有的欄位： 顏色標註、表徵、調查結果
    wsDest.Range(ColShSimpleAlertRange).value = ColShSimpleAlertName
    wsDest.Range(ColShSimpleRepRange).value = ColShSimpleRepName
    wsDest.Range(ColShSimpleResRange).value = ColShSimpleResName

    ' Check and highlight
    Dim cellTime As Date
    Dim strAcc   As String
    Dim i        As Long
    For i = RowDataBegin To lastRow
        ' set background color
        ' 每兩行標註為黃色底
        If (i - RowDataBegin) Mod 2 = 0 Then
            wsDest.Range(wsDest.Cells(i, ColShSimpleTSDate), wsDest.Cells(i, ColShSimpleAlertReason)).Interior.color = ColorYellow
        End If

        ' Color transations in the morning FIXME: move to viewer
        ' 凌晨交易
        If VarType(wsDest.Cells(i, ColShSimpleTSTime).value) = vbDouble Then
            cellTime = wsDest.Cells(i, ColShSimpleTSTime).value
            If ExcelUtils.IsMorningTime(cellTime) = True Then
                ExcelUtils.HighlightRow SheetNameSimple, i, ReasonTSInMorning, ColorBlue
            End If
        End If

        ' Check amount
        ' 檢查收入 支出，計算異常大額支出，並依照不同標準著色
        If (Not IsEmpty(wsDest.Cells(i, ColShSimpleTSOut)) Or Not IsEmpty(wsDest.Cells(i, ColShSimpleTSIn))) Then
            ' 支出大於 10 萬
            If IsNumeric(wsDest.Cells(i, ColShSimpleTSOut).value) And _
                         wsDest.Cells(i, ColShSimpleTSOut).value > ColShSimpleTSOut100K Then
                ' 支出大於 50 萬
                If wsDest.Cells(i, ColShSimpleTSOut).value > ColShSimpleTSOut500K Then
                    ExcelUtils.HighlightRow SheetNameSimple, i, ReasonTSLargeAmount, ColorPink
                Else
                    ExcelUtils.HighlightRow SheetNameSimple, i, ReasonTSLargeAmount, ColorOrange
                End If
            ' 收入大於 10 萬
            ElseIf IsNumeric(wsDest.Cells(i, ColShSimpleTSIn).value) And wsDest.Cells(i, ColShSimpleTSIn).value > ColShSimpleTSOut100K Then
                ' 收入大於 50 萬
                If wsDest.Cells(i, ColShSimpleTSIn).value > ColShSimpleTSOut500K Then
                    ExcelUtils.HighlightRow SheetNameSimple, i, ReasonTSLargeAmount, ColorPink
                Else
                    ExcelUtils.HighlightRow SheetNameSimple, i, ReasonTSLargeAmount, ColorOrange
                End If
            End If
        End If

        ' check blacklist
        ' 檢查警示戶帳號
        strAcc = wsDest.Cells(i, ColShSimpleTSInOutAcc)
        If Not IsEmpty(strAcc) And VarType(strAcc) = vbString Then
            If ExtraData.IsWarningAcc(strAcc) = True Then
                'Debug.Print " 警示戶　hit!!! " & strAcc
                ExcelUtils.HighlightRow SheetNameSimple, i, ReasonDWLAlert, ColorRed
            End If
        End If
    Next i

    ' FIXME: pass object
    ' 把交易日期相同的 merge起來
    MergeCellsByDate
End Sub

Sub MergeCellsByDate()
    Dim lastRow  As Long
    Dim currCell As Range
    Dim currDate As String
    Dim wsDest   As Worksheet
    Dim color    As Long

    color = 1
    Set wsDest = ThisWorkbook.Sheets(SheetNameSimple)

    ' Find the last row with data in column A
    lastRow = ExcelUtils.GetLastRowNumber(wsDest)

    ' Start from row 9 since the real data starts from there
    For Each currCell In wsDest.Range("A9:A" & lastRow)
        'Debug.Print "current: " & currDate & " now: " & currCell.Value
        If Trim(currCell.value) = currDate And Not IsEmpty(currCell.value) Then
            ' Merge the cells without warning
            wsDest.Range(currCell, currCell.Offset(-1, 0)).Merge
            wsDest.Range(currCell, currCell.Offset(-1, 0)).VerticalAlignment = xlCenter

            If (color) Mod 2 = 0 Then
                wsDest.Range(currCell, currCell.Offset(-1, 0)).Interior.color = ColorYellow
            Else
                wsDest.Range(currCell, currCell.Offset(-1, 0)).Interior.color = ColorWhite
            End If
        Else
            ' Update the current date
            currDate = Trim(currCell.value)
            Utils.CountOne color

            If (color) Mod 2 = 0 Then
                currCell.Interior.color = ColorYellow
            Else
                currCell.Interior.color = ColorWhite
            End If
        End If
    Next currCell
End Sub

' 這個子函式用來把 1原始資料 正規化成 2.2清整後資料
Private Sub NormalizeRawData()
    Dim wsSrc  As Worksheet
    Dim wsDest As Worksheet

    Set wsSrc = ThisWorkbook.Sheets(SheetNameOrginal)
    Set wsDest = ThisWorkbook.Sheets(SheetNameInputData)

    'Range(posRunStatus).value = BtnStatusStartRebase

    ExcelUtils.ClearSheet SheetNameInputData

    ' 抓客戶姓名
    If wsSrc.Range(ColShInDataBeginRange) = "" And _
       InStr(1, wsSrc.Range(ColShInDataCustomerRange), ColShInDataCustomerRegex) = 1 Then
        MainAccName = wsSrc.Range(ColShInDataCustNameRange)
    End If

    ' 抓客戶帳號
    If wsSrc.Range(ColShInDataBeginRange) = "" And _
        InStr(1, wsSrc.Range(ColShInDataAccRange), ColShInDataAccRegex) = 1 Then
        MainAccId = wsSrc.Range(ColShInDataAccIDRange)
    End If

    ' 如果 A1 是"交易日期"，表示原始資料不包含 客戶姓名 / 客戶帳號，
    ' 這種情況不管如何就往下寫， 這是work round，在還沒有填寫客戶姓名與帳號時假裝有。
    ' 這裡是做出 7 行空白，第 8行寫 header，真的內容從第 9 行開始
    If wsSrc.Range(ColShInDataBeginRange) = ColShInDataTSDateRegex Then
        wsSrc.Rows(RowShInDataEmpty).Insert Shift:=xlDown
    End If

    Dim i        As Long
    Dim lastRow  As Long
    Dim arrSrc() As Variant
    Dim arrDst() As Variant

    arrSrc = ExcelUtils.GetShOrgColsForRawData()
    arrDst = ExcelUtils.GetShInDataColsToFillRawData()

    lastRow = ExcelUtils.GetLastRowNumber(wsSrc)

    ExcelUtils.ApplySheetInputDataStyle

    ' Loop through each column ID in the array
    For i = 0 To UBound(arrSrc)
        wsDest.Range(wsDest.Range(arrDst(i) & "1"), _
                    wsDest.Range(arrDst(i) & CStr(lastRow))).value = _
                    wsSrc.Range(wsSrc.Range(arrSrc(i) & "1"), wsSrc.Range(arrSrc(i) & CStr(lastRow))).value
    Next i

    ' Prepare headers
    ' 寫header 的name
    Dim arrExtHs As Variant
    arrExtHs = ExcelUtils.GetShInDataExtraHeaderArray()
    For i = LBound(arrExtHs) To UBound(arrExtHs)
        If Len(Trim(arrExtHs(i))) > 0 Then
            wsDest.Cells(RowHeaderBegin, i).value = Trim(arrExtHs(i))
        End If
    Next i

    Dim wsATM          As Worksheet
    Dim dictATMCity    As Object
    Dim dictATMLoc     As Object
    Dim wsBranch       As Worksheet
    Dim dictBranchName As Object
    Dim dictBranchCity As Object

    Set wsATM = ThisWorkbook.Sheets(SheetNameATM)
    Set dictATMCity = ExtraData.GetKVPairsToDict(wsATM, "C", "M")
    Set dictATMLoc = ExtraData.GetKVPairsToDict(wsATM, "C", "E")

    Set wsBranch = ThisWorkbook.Sheets(SheetNameBranch)
    Set dictBranchName = ExtraData.GetKVPairsToDict(wsBranch, "A", "D")
    Set dictBranchCity = ExtraData.GetKVPairsToDict(wsBranch, "A", "F")

    Dim tmpstr, strTmp As String
    Dim amtWithdraw    As Long
    Dim amtDeposit     As Long
    Dim strSummary     As String
    Dim strSerial      As String
    Dim strNote        As String
    Dim strStore       As String
    Dim strClerk       As String
    Dim strTrans       As String
    Dim strAccount     As String
    Dim vAccName       As String
    Dim wAccName       As String
    Dim pAccName       As String

    For i = lastRow To RowDataBegin Step -1
        strStore = wsSrc.Cells(i, ColShOrgBranchID).value
        strClerk = wsSrc.Cells(i, ColShOrgTSTeller).value
        strSummary = wsSrc.Cells(i, ColShOrgSummary).value

        amtWithdraw = wsSrc.Cells(i, ColShOrgAmtDraw).value
        amtDeposit = wsSrc.Cells(i, ColShOrgAmtDeposit).value
        strAccount = wsSrc.Cells(i, ColShOrgAccount)
        strSerial = wsSrc.Cells(i, ColShOrgSerialCode).value
        strNote = wsSrc.Cells(i, ColShOrgChannel).value

        wsDest.Cells(i, ColShInDataTSMonth).value = Left(wsSrc.Cells(i, ColShOrgTSDate).Text, 7)
        wsDest.Cells(i, ColShInDataTSOClock).value = Left(wsSrc.Cells(i, ColShOrgTSTime).Text, 2) & "時"

        ' item 19, 2024/3/27, 保留手續費與利息後，以下暫時不會用到。
        'If ExcelUtils.IsSkipInSummary(strSummary) = True Then
        '    wsDest.Rows(i).EntireRow.Delete
        '    GoTo GoSkip
        'End If

        If amtWithdraw <> 0 Then
            wsDest.Cells(i, "H").value = amtWithdraw
        ElseIf amtDeposit <> 0 Then
            wsDest.Cells(i, "H").value = amtDeposit
        End If

        '"交易摘要"欄位
        wsDest.Cells(i, ColShInDataTSSummary).value = ExcelUtils.ConvToColTSSummary(strSummary)

        '"銀行代碼"欄位
        wsDest.Cells(i, ColShInDataBankCode).value = ExcelUtils.ConvToBankCode(strSerial)

        ' 交易通路
        wsDest.Cells(i, ColShInDataTSChannel).value = ExcelUtils.ConvToChannel(strNote, strStore, _
                                                        strClerk, strSummary, dictBranchName.Exists(strStore))

        ' "TranType"
        If strStore = "" Then
            wsDest.Cells(i, ColShInDataTSType).value = ""
        Else
            If ExcelUtils.IsAutoTSID(strStore) = True Then
                ' 自動化交易
                wsDest.Cells(i, ColShInDataTSType).value = ColShInDataChATMPostfix
            Else
                ' 非自動化通路
                If IsAutoClerk(strClerk) = False Then
                    If dictBranchName.Exists(strStore) Then
                        wsDest.Cells(i, ColShInDataTSType).value = ColShInDataChBRPostfix
                        wsDest.Cells(i, ColShInDataBranchName).value = dictBranchName(strStore) & "分行"
                    Else
                        wsDest.Cells(i, ColShInDataTSType).value = ""
                        wsDest.Cells(i, ColShInDataBranchName).value = ColShInDataChNBRPostfix & strStore
                    End If

                    If dictBranchCity.Exists(strStore) Then
                        wsDest.Cells(i, ColShInDataBranchCity).value = Left(dictBranchCity(strStore), 3)
                        wsDest.Cells(i, ColShInDataBranchArea).value = Mid(dictBranchCity(strStore), 4)
                        wsDest.Cells(i, ColShInDataTSLoc).value = dictBranchCity(strStore)
                    Else
                        wsDest.Cells(i, ColShInDataBranchCity).value = ""        ' wsOut.Cells(i, "F").Value FIXME:how to handle not found
                        wsDest.Cells(i, ColShInDataBranchArea).value = ""
                    End If
                End If

            End If
        End If

        ' ATM related
        ' 處理 2.2清整後資料裡的 ATM相關欄位
        If strClerk <> "" Then
            If dictATMLoc.Exists(strClerk) Then
                wsDest.Cells(i, ColShInDataATMLoc).value = dictATMLoc(strClerk)
            Else
                wsDest.Cells(i, ColShInDataATMLoc).value = ""
            End If

            If dictATMCity.Exists(strClerk) Then
                wsDest.Cells(i, ColShInDataATMCity).value = Left(dictATMCity(strClerk), 3)
                wsDest.Cells(i, ColShInDataATMArea).value = Mid(dictATMCity(strClerk), 4)
                wsDest.Cells(i, ColShInDataTSLoc).value = dictATMCity(strClerk)
            Else
                wsDest.Cells(i, ColShInDataATMCity).value = ""        ' wsOut.Cells(i, "F").Value FIXME:how to handle not found
                wsDest.Cells(i, ColShInDataATMArea).value = ""        '
            End If
        End If

        ' 重置三個變數，以便逐行取得正確值。
        vAccName = ""
        wAccName = ""
        pAccName = ""
        strTmp = ExtraData.ConvVAccName(wsDest.Cells(i, ColShInDataBankCode).value, strAccount)
        If strTmp <> "" Then
            Dim ansArr As Variant
            ansArr = Split(strTmp, GeneralDelimiter) ' ansArr(0) 是原先的 colAccount "901+ID(英文2碼+身分證字號後9碼) (812 虛擬帳號)"
                                                     ' ansArr(1) 是新的資訊 戶名 "繳卡費(一般信用卡)", "南山人壽保險股份有限公司" etc.
            wsDest.Cells(i, ColShInDataVAccReason).value = ansArr(0)
            wsDest.Cells(i, ColShInDataVAccCName).value = ansArr(1)

            vAccName = ansArr(1)
        End If

        If ExtraData.IsWarningAcc(strAccount) = True Then
            'Debug.Print " 警示戶　hit!!! " & strAccount
            wsDest.Cells(i, ColShInDataWAccCName).value = SheetNameBadAcc

            wAccName = SheetNameBadAcc
        End If

        ' 在 PivotTable  樞紐分析表 會用到的顯示帳戶名
        pAccName = ExcelUtils.ConvPivotAccShowName(vAccName, wAccName)
        wsDest.Cells(i, ColShInDataPAccCName).value = strAccount
        If Len(pAccName) > 0 Then
            wsDest.Cells(i, ColShInDataPAccCName).value = CStr(pAccName)
        End If
GoSkip:
    Next i

    wsDest.Range(ColShInDataRangePrefix & CStr(lastRow)).AutoFilter
End Sub

' 這個子函式用來畫出3.2金流與交易對手 的樞紐分析表，依序由上而下畫出
Sub CreatePivotTables()
    Dim wsSrc      As Worksheet
    Dim wsDest     As Worksheet
    Dim dataRange  As Range
    Dim rowCurrent As Long
    Dim rowPivot   As Long

    Dim i            As Long
    Dim pos          As Integer
    Dim pitems       As Object
    Dim colCityNames As Variant
    colCityNames = ExcelUtils.GetCityNameArray()

    ' 設定工作表
    Set wsSrc = ThisWorkbook.Sheets(SheetNameInputData)
    Set wsDest = ThisWorkbook.Sheets(SheetNameMoney)

    ExcelUtils.RemoveAllPivotTables(SheetNameMoney)

    ' 設定資料範圍
    lastRow = ExcelUtils.GetLastRowNumber(wsSrc)

    Set dataRange = wsSrc.Range(ColShInDataRangePrefix & CStr(lastRow))

    rowCurrent = PivotTableRowStart

    Dim pCache1, pCache2, pCache3, pCache4, pCache5, pCache6, pCache7, pCache8, pCache9, pCache10 As pivotCache
    Dim ptRange1, ptRange2, ptRange3, ptRange4, ptRange5, ptRange6, ptRange7, ptRange8, ptRange9, ptRange10 As Range
    Dim pTable1, pTable2, pTable3, pTable4, pTable5, pTable6, pTable7, pTable8, pTable9, pTable10 As PivotTable

    ' =========================================================
    ' MyPivotTable1, 存入交易
    ' 設定樞紐分析表的快取範圍, Cache Range
    Set pCache1 = PivotUtils.CreateCaches(dataRange)
    ' 設定樞紐分析表的位置
    Set ptRange1 = wsDest.Range("A" & CStr(rowCurrent))
    ' 在指定位置插入樞紐分析表
    'Set pTable1 = CreatePivotTable(pCache1, PivotTableName01, ptRange1)
    Set pTable1 = pCache1.CreatePivotTable(TableDestination:=ptRange1, tableName:=PivotTableName01)
    ' 在樞紐分析表中設定行、列、值
    With pTable1
        .CompactLayoutColumnHeader = ColShMoneyTSModelName
        .CompactLayoutRowHeader = ColShMoneyYYYYMMName

        .PivotFields(ColShMoneyTSMonthName).Orientation = xlRowField
        .PivotFields(ColShMoneyTSSummaryName).Orientation = xlColumnField
        .PivotFields(ColShMoneyDepositName).Orientation = xlDataField ' 這行不能少...? 太神奇

        With .PivotFields(ColShMoneyDepositName)
            .Orientation = xlDataField
            .Function = xlSum            ' 設定聚合函數
            .NumberFormat = NumberFormat ' 設定數值格式
        End With

        ' 這裡如果放到 別的地方去做，包含 call function，就會畫不出 "加總 - In2" 之類的欄位...
        ' Why??
        With wsDest.Range("A" & CStr(.TableRange1.Cells(1, 1).row - 1))
            .value = ColShMoneyTSDepositName
            .Font.color = ColorWhite
            .Font.bold = True
            .Interior.color = ColorGreen
        End With

        ' Move to next position
        rowCurrent = rowCurrent + .TableRange1.Rows.Count + PivotTableGap
    End With
    ' =========================================================

    ' =========================================================
    ' MyPivotTable2, 支出交易
    ' 設定樞紐分析表的快取範圍, Cache Range
    Set pCache2 = PivotUtils.CreateCaches(dataRange)
    ' 設定樞紐分析表的位置
    Set ptRange2 = wsDest.Range("A" & CStr(rowCurrent))
    ' 在指定位置插入樞紐分析表
    'Set pTable2 = CreatePivotTable(pCache2, PivotTableName02, ptRange2)
    Set pTable2 = pCache2.CreatePivotTable(TableDestination:=ptRange2, tableName:=PivotTableName02)
    ' 在樞紐分析表中設定行、列、值
    With pTable2
        .CompactLayoutColumnHeader = ColShMoneyTSModelName
        .CompactLayoutRowHeader = ColShMoneyYYYYMMName

        .PivotFields(ColShMoneyTSMonthName).Orientation = xlRowField
        .PivotFields(ColShMoneyTSSummaryName).Orientation = xlColumnField
        .PivotFields(ColShMoneyWithdrawName).Orientation = xlDataField

        With .PivotFields(ColShMoneyWithdrawName)
            .Orientation = xlDataField
            .Function = xlSum            ' 設定聚合函數
            .NumberFormat = NumberFormat ' 設定數值格式
        End With

        With wsDest.Range("A" & CStr(.TableRange1.Cells(1, 1).row - 1))
            .value = ColShMoneyTSWithdrawName
            .Font.color = ColorWhite
            .Font.bold = True
            .Interior.color = ColorGreen
        End With

        ' Move to next position
        rowCurrent = rowCurrent + .TableRange1.Rows.Count + PivotTableGap
    End With
    ' =================================================

    Dim lenTableLong  As Long
    Dim lenTableShort As Long

    ' =================================================
    ' MyPivotTable3, "轉帳存入交易By總流量"
    Dim cntRowTable3  As Long
    Set pCache3 = PivotUtils.CreateCaches(dataRange)
    ' 設定樞紐分析表的位置
    Set ptRange3 = wsDest.Range("A" & CStr(rowCurrent))
    ' 在指定位置插入樞紐分析表
    'Set pTable3 = CreatePivotTable(pCache3, PivotTableName03, ptRange3)
    Set pTable3 = pCache3.CreatePivotTable(TableDestination:=ptRange3, tableName:=PivotTableName03)
    ' 在樞紐分析表中設定行、列、值
    With pTable3
        .PivotFields(ColShMoneyTSSummaryName).Orientation = xlPageField
        ' item 22, 讓虛擬帳號帳戶名，或是"警示戶" 取代原本的account
        xAccName = ColShMoneyPivotAccountName 'ColShMoneyAccountName
        .PivotFields(xAccName).Orientation = xlRowField

        .AddDataField .PivotFields(ColShMoneyDepositName), ColShMoneyCountInName, xlCount
        .AddDataField .PivotFields(ColShMoneyDepositName), ColShMoneySumIn2Name, xlSum
        .AddDataField .PivotFields(ColShMoneyDepositName), ColShMoneyRatioIn3Name, xlSum

        .PivotFields(ColShMoneySumIn2Name).NumberFormat = NumberFormat
        .PivotFields(ColShMoneyRatioIn3Name).Calculation = xlPercentOfTotal
        .PivotFields(ColShMoneyRatioIn3Name).NumberFormat = "0%"
        .PivotFields(ColShMoneyTSSummaryName).ClearAllFilters

        ' 有這個 item 才用這個 item filter
        bExisted = PivotUtils.CheckPivotItemExisted(wsDest, PivotTableName03, ColShMoneyTSSummaryName, ColTSSummaryVal04)
        If bExisted = True Then
            .PivotFields(ColShMoneyTSSummaryName).CurrentPage = ColTSSummaryVal04
        End If

        .PivotFields(xAccName).PivotFilters.Add2 Type:=xlValueIsGreaterThan, _
                                                DataField:=.PivotFields(ColShMoneySumIn2Name), Value1:=0
        .PivotFields(xAccName).AutoSort xlDescending, ColShMoneySumIn2Name

        rowPivot = .TableRange1.Cells(1, 1).row
        wsDest.Range("E" & CStr(rowPivot + 1)).value = ColShMoneyCustomerName
        wsDest.Range("F" & CStr(rowPivot + 1)).value = ColShMoneyCustomerIDName

        With wsDest.Range("E" & CStr(rowPivot + 1) & ":F" & CStr(rowPivot + 1))
            .Font.color = ColorWhite
            .Font.bold = True
            .Interior.color = ColorGreen
        End With

        With wsDest.Range("E" & CStr(rowPivot + 2) & ":F" & CStr(rowPivot + 6))
            .Interior.color = ColorYellow
            .Borders.LineStyle = xlContinuous
            .Borders.color = ColorGreen
            .Borders.TintAndShade = 0
            .Borders.Weight = xlThin
        End With

        With wsDest.Range("A" & CStr(rowPivot - 1))
            .value = ColShMoneyTSInByTrafficName
            .Font.color = ColorWhite
            .Font.bold = True
            .Interior.color = ColorGreen
        End With

        wsDest.Range("B" & CStr(rowPivot - 1)).value = HintDoubleClick
        ' 此變數在這裡 Set 後，可供 工作表18 裡的操作使用，
        ' 請勿移除！
        Set HeaderPivotTable3 = wsDest.Range("A" & CStr(rowPivot - 1))

        ' Keep top k only
        ' 記錄起來，等等跟並排的 table4 比較長度，取最長的那個。
        With .TableRange1
            .Rows.Hidden = False

            cntRowTable3 = .Rows.Count
        End With
    End With
    ' =================================================

    ' =================================================
    ' MyPivotTable4, 轉帳支出交易By總流量
    Set pCache4 = PivotUtils.CreateCaches(dataRange)
    ' 設定樞紐分析表的位置
    Set ptRange4 = wsDest.Range("H" & CStr(rowCurrent))
    ' 在指定位置插入樞紐分析表
    'Set pTable4 = CreatePivotTable(pCache4, PivotTableName04, ptRange4)
    Set pTable4 = pCache4.CreatePivotTable(TableDestination:=ptRange4, tableName:=PivotTableName04)
    ' 在樞紐分析表中設定行、列、值
    With pTable4
        .PivotFields(ColShMoneyTSSummaryName).Orientation = xlPageField
        ' item 22, 讓虛擬帳號帳戶名，或是"警示戶" 取代原本的account
        xAccName = ColShMoneyPivotAccountName 'ColShMoneyAccountName
        .PivotFields(xAccName).Orientation = xlRowField

        .AddDataField .PivotFields(ColShMoneyWithdrawName), ColShMoneyCountOutName, xlCount
        .AddDataField .PivotFields(ColShMoneyWithdrawName), ColShMoneySumOut2Name, xlSum
        .AddDataField .PivotFields(ColShMoneyWithdrawName), ColShMoneyRatioOut3Name, xlSum

        .PivotFields(ColShMoneySumOut2Name).NumberFormat = NumberFormat
        .PivotFields(ColShMoneyRatioOut3Name).Calculation = xlPercentOfTotal
        .PivotFields(ColShMoneyRatioOut3Name).NumberFormat = "0%"
        .PivotFields(ColShMoneyTSSummaryName).ClearAllFilters

        ' 有這個 item 才用這個 item filter
        bExisted = PivotUtils.CheckPivotItemExisted(wsDest, PivotTableName04, ColShMoneyTSSummaryName, ColTSSummaryVal04)
        If bExisted = True Then
            .PivotFields(ColShMoneyTSSummaryName).CurrentPage = ColTSSummaryVal04
        End If

        .PivotFields(xAccName).PivotFilters.Add2 Type:=xlValueIsGreaterThan, _
                                                              DataField:=.PivotFields(ColShMoneySumOut2Name), Value1:=0
        .PivotFields(xAccName).AutoSort xlDescending, ColShMoneySumOut2Name

        rowPivot = .TableRange1.Cells(1, 1).row
        wsDest.Range("L" & CStr(rowPivot + 1)).value = ColShMoneyCustomerName
        wsDest.Range("M" & CStr(rowPivot + 1)).value = ColShMoneyCustomerIDName

        With wsDest.Range("L" & CStr(rowPivot + 1) & ":M" & CStr(rowPivot + 1))
            .Font.color = ColorWhite
            .Font.bold = True
            .Interior.color = ColorGreen
        End With

        With wsDest.Range("L" & CStr(rowPivot + 2) & ":M" & CStr(rowPivot + 6))
            .Interior.color = ColorYellow
            .Borders.LineStyle = xlContinuous
            .Borders.color = ColorGreen
            .Borders.TintAndShade = 0
            .Borders.Weight = xlThin
        End With

        With wsDest.Range("H" & CStr(rowPivot - 1))
            .value = ColShMoneyTSOutByTrafficName
            .Font.color = ColorWhite
            .Font.bold = True
            .Interior.color = ColorGreen
        End With

        wsDest.Range("I" & CStr(rowPivot - 1)).value = HintDoubleClick
        ' 此變數在這裡 Set 後，可供 工作表18 裡的操作使用，
        ' 請勿移除！
        Set HeaderPivotTable4 = wsDest.Range("H" & CStr(rowPivot - 1))

        ' Keep top k only 並排在右邊時使用
        With .TableRange1
            .Rows.Hidden = False

            ' 看 table3 跟 table4 哪個比較多列，取最多列的那個當作 hidding 的標準
            If cntRowTable3 >= .Rows.Count Then
                lenTableLong = cntRowTable3
                lenTableShort = .Rows.Count
            Else
                lenTableLong = .Rows.Count
                lenTableShort = cntRowTable3
            End If

            If lenTableLong > PivotTableCutShow Then
                ' 把超過 PivotTableCutShow 的 row 隱藏起來
                .Rows(CStr(PivotTableCutShow) & ":" & CStr(lenTableLong - 1)).Hidden = True
                .Rows(CStr(lenTableShort)).Hidden = False ' unhide sum row of smaller table
            End If
        End With

        '' Move to next position
        rowCurrent = rowCurrent + lenTableLong + PivotTableGap + 3
    End With
    ' =================================================

    ' =================================================
    ' MyPivotTable5, 跨行轉帳存入交易By總流量
    Dim cntRowTable5 As Long
    Set pCache5 = PivotUtils.CreateCaches(dataRange)
    ' 設定樞紐分析表的位置
    Set ptRange5 = wsDest.Range("A" & CStr(rowCurrent))
    ' 在指定位置插入樞紐分析表
    'Set pTable5 = CreatePivotTable(pCache5, PivotTableName05, ptRange5)
    Set pTable5 = pCache5.CreatePivotTable(TableDestination:=ptRange5, tableName:=PivotTableName05)
    ' 在樞紐分析表中設定行、列、值
    With pTable5
        .PivotFields(ColShMoneyTSSummaryName).Orientation = xlPageField
        ' item 22, 讓虛擬帳號帳戶名，或是"警示戶" 取代原本的account
        xAccName = ColShMoneyPivotAccountName 'ColShMoneyAccountName
        .PivotFields(xAccName).Orientation = xlRowField

        .AddDataField .PivotFields(ColShMoneyDepositName), ColShMoneyCountInName, xlCount
        .AddDataField .PivotFields(ColShMoneyDepositName), ColShMoneySumIn2Name, xlSum
        .AddDataField .PivotFields(ColShMoneyDepositName), ColShMoneyRatioIn3Name, xlSum

        .PivotFields(ColShMoneySumIn2Name).NumberFormat = NumberFormat
        .PivotFields(ColShMoneyRatioIn3Name).Calculation = xlPercentOfTotal
        .PivotFields(ColShMoneyRatioIn3Name).NumberFormat = "0%"
        .PivotFields(ColShMoneyTSSummaryName).ClearAllFilters

        ' 有這個 item 才用這個 item filter
        bExisted = PivotUtils.CheckPivotItemExisted(wsDest, PivotTableName05, ColShMoneyTSSummaryName, ColTSSummaryVal03)
        If bExisted = True Then
            .PivotFields(ColShMoneyTSSummaryName).CurrentPage = ColTSSummaryVal03
        End If

        .PivotFields(xAccName).PivotFilters.Add2 Type:=xlValueDoesNotEqual, _
                                                              DataField:=.PivotFields(ColShMoneySumIn2Name), Value1:=0
        .PivotFields(xAccName).AutoSort xlDescending, ColShMoneySumIn2Name

        rowPivot = .TableRange1.Cells(1, 1).row
        wsDest.Range("E" & CStr(rowPivot + 1)).value = ColShMoneyCustomerName
        wsDest.Range("F" & CStr(rowPivot + 1)).value = ColShMoneyCustomerIDName

        With wsDest.Range("E" & CStr(rowPivot + 1) & ":F" & CStr(rowPivot + 1))
            .Font.color = ColorWhite
            .Font.bold = True
            .Interior.color = ColorGreen
        End With

        With wsDest.Range("E" & CStr(rowPivot + 2) & ":F" & CStr(rowPivot + 6))
            .Interior.color = ColorYellow
            .Borders.LineStyle = xlContinuous
            .Borders.color = ColorGreen
            .Borders.TintAndShade = 0
            .Borders.Weight = xlThin
        End With

        With wsDest.Range("A" & CStr(rowPivot - 1))
            .value = ColShMoneyTSBrInByTrafficName
            .Font.color = ColorWhite
            .Font.bold = True
            .Interior.color = ColorGreen
        End With

        wsDest.Range("B" & CStr(rowPivot - 1)).value = HintDoubleClick
        ' 此變數在這裡 Set 後，可供 工作表18 裡的操作使用，
        ' 請勿移除！
        Set HeaderPivotTable5 = wsDest.Range("A" & CStr(rowPivot - 1))

        ' Keep top k only 並排在左邊使用
        With .TableRange1
            .Rows.Hidden = False

            cntRowTable5 = .Rows.Count
        End With
    End With
    ' =================================================

    ' =================================================
    ' MyPivotTable6, 跨行轉帳支出交易By總流量
    Set pCache6 = PivotUtils.CreateCaches(dataRange)
    ' 設定樞紐分析表的位置
    Set ptRange6 = wsDest.Range("H" & CStr(rowCurrent))
    ' 在指定位置插入樞紐分析表
    'Set pTable6 = CreatePivotTable(pCache6, PivotTableName06, ptRange6)
    Set pTable6 = pCache6.CreatePivotTable(TableDestination:=ptRange6, tableName:=PivotTableName06)
    ' 在樞紐分析表中設定行、列、值
    With pTable6
        .PivotFields(ColShMoneyTSSummaryName).Orientation = xlPageField
        ' item 22, 讓虛擬帳號帳戶名，或是"警示戶" 取代原本的account
        xAccName = ColShMoneyPivotAccountName 'ColShMoneyAccountName
        .PivotFields(xAccName).Orientation = xlRowField

        .AddDataField .PivotFields(ColShMoneyWithdrawName), ColShMoneyCountOutName, xlCount
        .AddDataField .PivotFields(ColShMoneyWithdrawName), ColShMoneySumOut2Name, xlSum
        .AddDataField .PivotFields(ColShMoneyWithdrawName), ColShMoneyRatioOut3Name, xlSum

        .PivotFields(ColShMoneySumOut2Name).NumberFormat = NumberFormat
        .PivotFields(ColShMoneyRatioOut3Name).Calculation = xlPercentOfTotal
        .PivotFields(ColShMoneyRatioOut3Name).NumberFormat = "0%"
        .PivotFields(ColShMoneyTSSummaryName).ClearAllFilters

        ' 有這個 item 才用這個 item filter
        bExisted = PivotUtils.CheckPivotItemExisted(wsDest, PivotTableName06, ColShMoneyTSSummaryName, ColTSSummaryVal03)
        If bExisted = True Then
            .PivotFields(ColShMoneyTSSummaryName).CurrentPage = ColTSSummaryVal03
        End If

        .PivotFields(xAccName).PivotFilters.Add2 Type:=xlValueDoesNotEqual, _
                                                              DataField:=.PivotFields(ColShMoneySumOut2Name), Value1:=0
        .PivotFields(xAccName).AutoSort xlDescending, ColShMoneySumOut2Name

        rowPivot = .TableRange1.Cells(1, 1).row
        wsDest.Range("L" & CStr(rowPivot + 1)).value = ColShMoneyCustomerName
        wsDest.Range("M" & CStr(rowPivot + 1)).value = ColShMoneyCustomerIDName

        With wsDest.Range("L" & CStr(rowPivot + 1) & ":M" & CStr(rowPivot + 1))
            .Font.color = ColorWhite
            .Font.bold = True
            .Interior.color = ColorGreen
        End With

        With wsDest.Range("L" & CStr(rowPivot + 2) & ":M" & CStr(rowPivot + 6))
            .Interior.color = ColorYellow
            .Borders.LineStyle = xlContinuous
            .Borders.color = ColorGreen
            .Borders.TintAndShade = 0
            .Borders.Weight = xlThin
        End With

        With wsDest.Range("H" & CStr(rowPivot - 1))
            .value = ColShMoneyTSBrOutByTrafficName
            .Font.color = ColorWhite
            .Font.bold = True
            .Interior.color = ColorGreen
        End With

        wsDest.Range("I" & CStr(rowPivot - 1)).value = HintDoubleClick
        ' 此變數在這裡 Set 後，可供 工作表18 裡的操作使用，
        ' 請勿移除！
        Set HeaderPivotTable6 = wsDest.Range("H" & CStr(rowPivot - 1))

        ' Keep top k only
        With .TableRange1
            .Rows.Hidden = False

            If cntRowTable5 >= .Rows.Count Then
                lenTableLong = cntRowTable5
                lenTableShort = .Rows.Count
            Else
                lenTableLong = .Rows.Count
                lenTableShort = cntRowTable5
            End If

            If lenTableLong > PivotTableCutShow Then
                ' 把超過 PivotTableCutShow 的 row 隱藏起來
                .Rows(CStr(PivotTableCutShow) & ":" & CStr(lenTableLong - 1)).Hidden = True
                .Rows(CStr(lenTableShort)).Hidden = False
            End If
        End With

        ' Move to next position
        rowCurrent = rowCurrent + lenTableLong + PivotTableGap + 3 '6
    End With
    ' =================================================

    ' ATM related PivotTables
    If GotATM > 0 Then
        ' =================================================
        ' MyPivotTable7, ATM存入交易時間
        Set pCache7 = PivotUtils.CreateCaches(dataRange)
        ' 設定樞紐分析表的位置
        Set ptRange7 = wsDest.Range("A" & CStr(rowCurrent))
        ' 在指定位置插入樞紐分析表
        'Set pTable7 = CreatePivotTable(pCache7, PivotTableName07, ptRange7)
        Set pTable7 = pCache7.CreatePivotTable(TableDestination:=ptRange7, tableName:=PivotTableName07)
        ' 在樞紐分析表中設定行、列、值
        With pTable7
            .CompactLayoutColumnHeader = ColShMoneyTimeName
            .CompactLayoutRowHeader = ColShMoneyYYYYMMName

            .PivotFields(ColShMoneyATMCityName).Orientation = xlPageField
            .PivotFields(ColShMoneyTSSummaryName).Orientation = xlPageField
            .PivotFields(ColShMoneyTSMonthName).Orientation = xlRowField
            .PivotFields(ColShMoneyDepositName).Orientation = xlDataField
            .PivotFields(ColShMoneyTSOClockName).Orientation = xlColumnField
            .PivotFields(ColShMoneyTSSummaryName).ClearAllFilters

            ' 有這個 item 才用這個 item filter
            bExisted = PivotUtils.CheckPivotItemExisted(wsDest, PivotTableName07, ColShMoneyTSSummaryName, ColTSSummaryVal01)
            If bExisted = True Then
                .PivotFields(ColShMoneyTSSummaryName).CurrentPage = ColTSSummaryVal01
            End If

            With wsDest.Range("A" & CStr(.TableRange1.Cells(1, 1).row - 1))
                .value = ColShMoneyATMDepositTimeName
                .Font.color = ColorWhite
                .Font.bold = True
                .Interior.color = ColorGreen
            End With

            ' Move to next position
            rowCurrent = rowCurrent + .TableRange1.Rows.Count + PivotTableGap
        End With
        ' =================================================

        ' =================================================
        ' MyPivotTable8, ATM存入交易地點
        Set pCache8 = PivotUtils.CreateCaches(dataRange)
        ' 設定樞紐分析表的位置
        Set ptRange8 = wsDest.Range("A" & CStr(rowCurrent))
        ' 在指定位置插入樞紐分析表
        'Set pTable8 = CreatePivotTable(pCache8, PivotTableName08, ptRange8)
        Set pTable8 = pCache8.CreatePivotTable(TableDestination:=ptRange8, tableName:=PivotTableName08)
        ' 在樞紐分析表中設定行、列、值
        With pTable8
            .CompactLayoutColumnHeader = ColShMoneyCityName
            .CompactLayoutRowHeader = ColShMoneyYYYYMMName

            .PivotFields(ColShMoneyATMCityName).Orientation = xlColumnField
            .PivotFields(ColShMoneyTSSummaryName).Orientation = xlPageField
            .PivotFields(ColShMoneyTSMonthName).Orientation = xlRowField
            .PivotFields(ColShMoneyDepositName).Orientation = xlDataField
            .PivotFields(ColShMoneyTSSummaryName).ClearAllFilters

            ' 有這個 item 才用這個 item filter
            bExisted = PivotUtils.CheckPivotItemExisted(wsDest, PivotTableName08, ColShMoneyTSSummaryName, ColTSSummaryVal01)
            If bExisted = True Then
                .PivotFields(ColShMoneyTSSummaryName).CurrentPage = ColTSSummaryVal01
            End If

            Set pitems = .PivotFields(ColShMoneyATMCityName).PivotItems
            pos = 2
            If pitems.Count > 0 And pitems(1).name = "" Then
                pitems(1).name = ColATMValOt
            End If

            For i = LBound(colCityNames) To UBound(colCityNames)
                If Utils.ObjectContainsItem(pitems, colCityNames(i)) = True Then
                    pitems(colCityNames(i)).position = pos
                    pos = pos + 1
                End If
            Next i

            With wsDest.Range("A" & CStr(.TableRange1.Cells(1, 1).row - 1))
                .value = ColShMoneyATMDepositLocName
                .Font.color = ColorWhite
                .Font.bold = True
                .Interior.color = ColorGreen
            End With

            ' Move to next position
            rowCurrent = rowCurrent + .TableRange1.Rows.Count + PivotTableGap + 3
        End With
        ' =================================================

        ' =================================================
        ' MyPivotTable9, ATM領出交易時間
        Set pCache9 = PivotUtils.CreateCaches(dataRange)
        ' 設定樞紐分析表的位置
        Set ptRange9 = wsDest.Range("A" & CStr(rowCurrent))
        ' 在指定位置插入樞紐分析表
        'Set pTable9 = CreatePivotTable(pCache9, PivotTableName09, ptRange9)
        Set pTable9 = pCache9.CreatePivotTable(TableDestination:=ptRange9, tableName:=PivotTableName09)
        ' 在樞紐分析表中設定行、列、值
        With pTable9
            .CompactLayoutColumnHeader = ColShMoneyTimeName
            .CompactLayoutRowHeader = ColShMoneyYYYYMMName

            .PivotFields(ColShMoneyATMCityName).Orientation = xlPageField
            .PivotFields(ColShMoneyTSSummaryName).Orientation = xlPageField
            .PivotFields(ColShMoneyTSMonthName).Orientation = xlRowField
            .PivotFields(ColShMoneyWithdrawName).Orientation = xlDataField
            .PivotFields(ColShMoneyTSOClockName).Orientation = xlColumnField
            .PivotFields(ColShMoneyTSSummaryName).ClearAllFilters

            ' 有這個 item 才用這個 item filter
            bExisted = PivotUtils.CheckPivotItemExisted(wsDest, PivotTableName09, ColShMoneyTSSummaryName, ColTSSummaryVal01)
            If bExisted = True Then
                .PivotFields(ColShMoneyTSSummaryName).CurrentPage = ColTSSummaryVal01
            End If

            With wsDest.Range("A" & CStr(.TableRange1.Cells(1, 1).row - 1))
                .value = ColShMoneyATMWithdrawTimeName
                .Font.color = ColorWhite
                .Font.bold = True
                .Interior.color = ColorGreen
            End With

            ' Move to next position
            rowCurrent = rowCurrent + .TableRange1.Rows.Count + PivotTableGap
        End With
        ' =================================================

        ' =================================================
        ' MyPivotTable10, ATM領出交易地點
        Set pCache10 = PivotUtils.CreateCaches(dataRange)
        ' 設定樞紐分析表的位置
        Set ptRange10 = wsDest.Range("A" & CStr(rowCurrent))
        ' 在指定位置插入樞紐分析表
        'Set pTable10 = CreatePivotTable(pCache10, PivotTableName10, ptRange10)
        Set pTable10 = pCache10.CreatePivotTable(TableDestination:=ptRange10, tableName:=PivotTableName10)
        ' 在樞紐分析表中設定行、列、值
        With pTable10
            .CompactLayoutColumnHeader = ColShMoneyCityName
            .CompactLayoutRowHeader = ColShMoneyYYYYMMName

            .PivotFields(ColShMoneyATMCityName).Orientation = xlColumnField
            .PivotFields(ColShMoneyTSSummaryName).Orientation = xlPageField
            .PivotFields(ColShMoneyTSMonthName).Orientation = xlRowField
            .PivotFields(ColShMoneyWithdrawName).Orientation = xlDataField
            .PivotFields(ColShMoneyTSSummaryName).ClearAllFilters

            ' 有這個 item 才用這個 item filter
            bExisted = PivotUtils.CheckPivotItemExisted(wsDest, PivotTableName10, ColShMoneyTSSummaryName, ColTSSummaryVal01)
            If bExisted = True Then
                .PivotFields(ColShMoneyTSSummaryName).CurrentPage = ColTSSummaryVal01
            End If

            Set pitems = .PivotFields(ColShMoneyATMCityName).PivotItems
            pos = 2
            If pitems.Count > 0 And pitems(1).name = "" Then
                pitems(1).name = ColATMValOt
            End If

            For i = LBound(colCityNames) To UBound(colCityNames)
                If Utils.ObjectContainsItem(pitems, colCityNames(i)) = True Then
                    pitems(colCityNames(i)).position = pos
                    pos = pos + 1
                End If
            Next i

            With wsDest.Range("A" & CStr(.TableRange1.Cells(1, 1).row - 1))
                .value = ColShMoneyATMWithdrawLocName
                .Font.color = ColorWhite
                .Font.bold = True
                .Interior.color = ColorGreen
            End With

            ' Move to next position
            rowCurrent = rowCurrent + .TableRange1.Rows.Count + PivotTableGap
        End With
        ' =================================================
    End If
    ' End of ATM related PivotTables

    ExcelUtils.ApplySheetMoneyStyle
End Sub
