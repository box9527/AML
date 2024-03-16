
Private Const idCountAlertMonth As Integer = 1          ' 找出 "較之前多3~5倍 超出閥值的月份 總次數", 這裡定義的是 taskID
Private Const idAmountAlertMonth As Integer = 2         ' 找出 "較之前多3~5倍 超出閥值的月份 總金額", 這裡定義的是 taskID
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

' 拉出來變成 Global Variables 用來控制 交易對手 裡要不要秀 ATM 的 PivotTable
Private gCntMultiATM As Integer
Private gCntMultiDeposit As Integer

Function LookupBranchInfo(branchID As Variant) As Variant
    Dim ws          As Worksheet
    Dim lastRow     As Long
    Dim branchRange As Range
    Dim branchCell  As Range

    Set ws = ThisWorkbook.Sheets(SheetNameBranch)

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row

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
    Else
        ' Return an error value Or handle the Case when branch ID is Not found
        LookupBranchInfo = CVErr(xlErrValue)
    End If
End Function

Function FindMonthsExceedingThreshold(dateRow As Integer, amountRow As Integer, taskID As Integer) As Variant
    Dim ws          As Worksheet
    Dim lastColumn  As Long
    Dim columnIndices() As Long
    Dim i           As Long

    Set ws = Worksheets(SheetNameIntermediate)
    ' Find the last column With data in the specified rows
    lastColumn = ws.Cells(dateRow, ws.columns.Count).End(xlToLeft).Column

    ' Initialize array To store column indices
    ReDim columnIndices(1 To 1)

    columnIndices(1) = EmptyArrayValue        ' Note:　this is a special marker for empty array

    If lastColumn < 2 Then
        FindMonthsExceedingThreshold = columnIndices
        Exit Function
    End If

    Dim col         As Long
    Dim currentMonthAmount As Double
    Dim previousMonthAmount As Double
    i = 1
    ' Loop through the columns To identify months
    For col = 2 To lastColumn        ' Assuming the first column contains headers
        ' Get the amount values For the current And previous months
        currentMonthAmount = CDbl(ws.Cells(amountRow, col).Value)
        previousMonthAmount = CDbl(ws.Cells(amountRow, col - 1).Value)
        ' Check If the current month's amount is three times larger than the previous month
        If taskID = idAmountAlertMonth Then
            'FIXME: 這個閥值應該參數化
            If currentMonthAmount >= 3 * previousMonthAmount Then
                ' Add the column index To the array
                ReDim Preserve columnIndices(1 To i)
                columnIndices(i) = col
                i = i + 1
            End If
        ElseIf taskID = idCountAlertMonth Then
            If currentMonthAmount >= 3 * previousMonthAmount Then
                ' Add the column index To the array
                ReDim Preserve columnIndices(1 To i)
                columnIndices(i) = col
                i = i + 1
            End If
        ElseIf taskID = idOver1MillionAlertMonth Then
            If currentMonthAmount >= (previousMonthAmount + 1000000) Then
                ' Add the column index To the array
                ReDim Preserve columnIndices(1 To i)
                columnIndices(i) = col
                i = i + 1
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
    Dim ws          As Worksheet
    Set ws = ThisWorkbook.Worksheets(SheetNameInputData)
    Dim wsInt       As Worksheet
    Set wsInt = ThisWorkbook.Worksheets(SheetNameIntermediate)

    Dim lastRow     As Long
    Dim currentGroup As Range
    Dim row         As Long
    Dim idxcolOutput As Long
    Dim currentYear, currentMonth As Integer

    idxcolOutput = 1        ' output column index

    lastRow = ws.Cells(ws.Rows.Count, ColTSDate).End(xlUp).row
    ' Loop through rows starting from row rowDataBegin, day by day
    For row = RowDataBegin To lastRow
        Dim dateValue As Date
        Dim y, m    As Integer

        dateValue = ws.Cells(row, ColTSDate).Value
        y = year(dateValue)
        m = month(dateValue)
        ' Check If the date has changed Or If it's the first row
        If row = RowDataBegin Or y <> currentYear Or m <> currentMonth Then
            ' Process the current group of rows
            If Not currentGroup Is Nothing Then
                ProcessGroupByMonth currentGroup, CStr(currentYear) & "/" & CStr(currentMonth), idxcolOutput
                idxcolOutput = idxcolOutput + 1

                Set currentGroup = Nothing
            End If

            ' Initialize a New group For the current date
            Set currentGroup = ws.Range(ws.Cells(row, 1), ws.Cells(row, MaxSrcCol))

            currentYear = y
            currentMonth = m
        Else
            ' Add the row To the current group
            Set currentGroup = Union(currentGroup, ws.Range(ws.Cells(row, 1), ws.Cells(row, MaxSrcCol)))

        End If
    Next row

    ' Process the last group of rows
    If Not currentGroup Is Nothing Then
        ProcessGroupByMonth currentGroup, CStr(currentYear) & "/" & CStr(currentMonth), idxcolOutput
        idxcolOutput = idxcolOutput + 1

    End If

    ' 畫圖
    CreateBarChart RowTask1TotalCountXAxis, RowTask1TotalCountYAxis, Range("H3:O5"), "分月交易總次數", "每月次數"

    CreateBarChart RowTask23TotalAmountXAxis, RowTask23TotalAmountYAxis, Range("Q3:X5"), "分月交易總金額", "每月金額"

    ' Add markers For selected one
    Dim resultArray() As Long        ' 有超過閥值的月份
    Dim i           As Integer
    Dim strSummary  As String

    resultArray = FindMonthsExceedingThreshold(RowTask1TotalCountXAxis, RowTask1TotalCountYAxis, idCountAlertMonth)

    ' 寫文字摘要
    If IsThresholdArrayEmpty(resultArray) = False Then
        strSummary = "在分月交易數據中，發現在"
        For i = LBound(resultArray) To UBound(resultArray)
            If resultArray(i) > 0 Then
                strSummary = strSummary & CStr(wsInt.Cells(RowTask1TotalCountXAxis, resultArray(i)).Value) & " "
            End If
        Next i
        strSummary = strSummary & "有異常的月總筆數增幅"
        Range(posTask1Result).Font.color = ColorRed
        Range(posTask1Result).WrapText = True
        Range(posTask1Result).Value = strSummary
    End If

    resultArray = FindMonthsExceedingThreshold(RowTask23TotalAmountXAxis, RowTask23TotalAmountYAxis, idAmountAlertMonth)

    ' 寫文字摘要
    If IsThresholdArrayEmpty(resultArray) = False Then
        strSummary = "在分月交易數據中，發現在"
        For i = LBound(resultArray) To UBound(resultArray)
            If resultArray(i) > 0 Then
                strSummary = strSummary & CStr(wsInt.Cells(RowTask23TotalAmountXAxis, resultArray(i)).Value) & " "
            End If
        Next i
        strSummary = strSummary & "有異常的月總金額增幅"
        Range(posTask2Result).Font.color = ColorRed
        Range(posTask2Result).WrapText = True
        Range(posTask2Result).Value = strSummary
    End If

    resultArray = FindMonthsExceedingThreshold(RowTask23TotalAmountXAxis, RowTask23TotalAmountYAxis, idOver1MillionAlertMonth)

    ' 寫文字摘要
    If IsThresholdArrayEmpty(resultArray) = False Then
        strSummary = "在分月交易數據中，發現在"
        For i = LBound(resultArray) To UBound(resultArray)
            If resultArray(i) > 0 Then
                strSummary = strSummary & CStr(wsInt.Cells(RowTask23TotalAmountXAxis, resultArray(i)).Value) & " "
            End If
        Next i
        strSummary = strSummary & "有異常的月總金額增幅"
        Range(posTask3Result).Font.color = ColorRed
        Range(posTask3Result).WrapText = True
        Range(posTask3Result).Value = strSummary
    End If

End Sub

Sub ProcessGroupByMonth(groupRange As Range, colName As String, idxOut As Long)
    Dim intResult   As Long
    ' groupRange represents all rows With the same date
    intResult = groupRange.row + groupRange.Rows.Count - 1

    Dim sum         As Double
    Dim cell        As Range

    For Each cell In groupRange.columns(ColAmountTransfer).Cells
        If IsNumeric(cell.Value) Then
            sum = sum + CDbl(cell.Value)
        End If
    Next cell

    Dim wsInt       As Worksheet
    Set wsInt = ThisWorkbook.Worksheets(SheetNameIntermediate)

    wsInt.Cells(RowTask23TotalAmountXAxis, idxOut).Value = colName
    wsInt.Cells(RowTask23TotalAmountYAxis, idxOut).Value = sum

    wsInt.Cells(RowTask1TotalCountXAxis, idxOut).Value = colName
    wsInt.Cells(RowTask1TotalCountYAxis, idxOut).Value = groupRange.Rows.Count

End Sub

' TASK
Sub ProcessRowsByDay()
    Dim ws          As Worksheet
    Set ws = ThisWorkbook.Worksheets(SheetNameInputData)
    Dim wsInt       As Worksheet
    Set wsInt = ThisWorkbook.Worksheets(SheetNameIntermediate)

    Dim lastRow     As Long
    Dim currentGroup As Range
    Dim row         As Long
    Dim idxcolOutput As Long
    Dim idxcolSmallAmount As Long
    Dim idxcolDiff  As Long
    Dim currentYear, currentMonth, currentDay As Integer

    idxcolOutput = 1        ' output column index
    idxcolSmallAmount = 1        ' another output column
    idxcolDiff = 1
    lastRow = ws.Cells(ws.Rows.Count, ColTSDate).End(xlUp).row
    ' Loop through rows starting from row rowDataBegin, day by day
    For row = RowDataBegin To lastRow        ' lastRow FIXME: For debug only
        Dim dateValue As Date
        Dim y, m, d As Integer
        Dim ret     As Boolean

        dateValue = ws.Cells(row, ColTSDate).Value
        y = year(dateValue)
        m = month(dateValue)
        d = Day(dateValue)
        ' Check If the date has changed Or If it's the first row
        If row = RowDataBegin Or y <> currentYear Or m <> currentMonth Or d <> currentDay Then
            ' Process the current group of rows
            If Not currentGroup Is Nothing Then
                ret = ProcessGroupByDay(currentGroup, CStr(currentYear) & "/" & CStr(currentMonth) & "/" & CStr(currentDay), idxcolOutput)
                If ret = True Then
                    idxcolOutput = idxcolOutput + 1
                End If
                ret = ProcessGroupByDaySmallAmount(currentGroup, CStr(currentYear) & "/" & CStr(currentMonth) & "/" & CStr(currentDay), idxcolSmallAmount)
                If ret = True Then
                    idxcolSmallAmount = idxcolSmallAmount + 1
                End If

                ProcessGroupByDayDiff currentGroup, CStr(currentYear) & "/" & CStr(currentMonth) & "/" & CStr(currentDay), idxcolDiff
                idxcolDiff = idxcolDiff + 1

                Set currentGroup = Nothing
            End If

            ' Initialize a New group For the current date
            Set currentGroup = ws.Range(ws.Cells(row, 1), ws.Cells(row, MaxSrcCol))

            currentYear = y
            currentMonth = m
            currentDay = d
        Else
            ' Add the row To the current group
            Set currentGroup = Union(currentGroup, ws.Range(ws.Cells(row, 1), ws.Cells(row, MaxSrcCol)))

        End If
    Next row
    ' Process the last group of rows
    If Not currentGroup Is Nothing Then
        ProcessGroupByDay currentGroup, CStr(currentYear) & "/" & CStr(currentMonth) & "/" & CStr(currentDay), idxcolOutput

        ProcessGroupByDaySmallAmount currentGroup, CStr(currentYear) & "/" & CStr(currentMonth) & "/" & CStr(currentDay), idxcolSmallAmount

        ProcessGroupByDayDiff currentGroup, CStr(currentYear) & "/" & CStr(currentMonth) & "/" & CStr(currentDay), idxcolDiff

    End If

    Dim lastColumn  As Long
    Dim i           As Long
    Dim nDays       As Integer
    '  For task 11
    lastColumn = wsInt.Cells(RowTask12SmallTransferXAxis, wsInt.columns.Count).End(xlToLeft).Column

    nDays = 0
    For i = 1 To lastColumn
        If wsInt.Cells(RowTask12SmallTransferOneDollar, i).Value >= 1 Then
            nDays = nDays + 1
        End If
    Next i

    Range(posTask12Result).WrapText = True
    If nDays > 0 Then
        Range(posTask12Result).Value = "在觀察期間交易數據中，有 " & CStr(nDays) & " 日發現一元轉帳"
    End If

    CreateBarChart RowTask12SmallTransferXAxis, RowTask12SmallTransferOneDollar, Range("H18:O19"), "一元交易", "交易筆數"

    nDays = 0
    For i = 1 To lastColumn
        If wsInt.Cells(RowTask12SmallTransferHundred, i).Value >= 1 Then
            nDays = nDays + 1
        End If
    Next i
    Range(posTask13Result).WrapText = True
    If nDays > 0 Then
        Range(posTask13Result).Value = "在觀察期間交易數據中，有 " & CStr(nDays) & " 日發現小於百元轉帳"
    End If
    CreateBarChart RowTask12SmallTransferXAxis, RowTask12SmallTransferHundred, Range("Q18:X19"), "小於百元交易", "交易筆數"

    nDays = 0
    For i = 1 To lastColumn
        If wsInt.Cells(RowTask12SmallTransferMany, i).Value >= 1 Then
            nDays = nDays + 1
        End If
    Next i
    Range(posTask14Result).WrapText = True
    If nDays > 0 Then
        Range(posTask14Result).Value = "在觀察期間交易數據中，有 " & CStr(nDays) & " 日發現跨行交易筆數高案例"
    End If
    CreateBarChart RowTask12SmallTransferXAxis, RowTask12SmallTransferMany, Range("H20:X20"), "跨行交易", "交易筆數"

    ' For task 45

    Dim nTran       As Long
    Dim strSusDates As String
    Dim hasSusDates As Boolean

    strSusDates = ""
    hasSusDates = False

    Dim dict        As New Scripting.Dictionary

    lastColumn = wsInt.Cells(RowTask4WithdrawOverCounterXAxis, wsInt.columns.Count).End(xlToLeft).Column

    For i = 1 To lastColumn
        Dim val, key As Variant
        key = wsInt.Cells(RowTask4WithdrawOverCounterXAxis, i).Value
        val = wsInt.Cells(RowTask4WithdrawOverCounterYAxis, i).Value

        If val > 1 Then
            strSusDates = strSusDates & " " & key
            hasSusDates = True
        End If

        If Not dict.Exists(val) Then
            dict.Add val, 1
        Else
            dict(val) = dict(val) + 1
        End If
    Next i

    Dim col         As Long
    Dim cntTotal    As Long

    col = 1
    For i = 1 To lastColumn
        If dict.Exists(i) Then
            wsInt.Cells(RowTask5WithdrawSummaryXAxis, col).Value = "發生" & CStr(i) & "筆"
            wsInt.Cells(RowTask5WithdrawSummaryYAxis, col).Value = dict(i)
            col = col + 1
            cntTotal = cntTotal + dict(i)
        End If
    Next i

    ' Task 4
    CreateBarChart RowTask5WithdrawSummaryXAxis, RowTask5WithdrawSummaryYAxis, Range("H7:O8"), "可疑大額提領", "天數"

    ' 寫摘要
    If hasSusDates = False Then
        Range(posTask4Result).Value = "在當日臨櫃提領現金交易數據中，沒有發現異常"
    End If

    ' 多間分行提領

    ' Specify the row index you want To traverse
    Dim idxCol      As Integer
    Dim hasMulti    As Boolean
    Dim strHits     As String
    Dim cellKey     As String
    Dim cellValue   As String

    hasMulti = False
    strHits = ""

    ' Loop through each cell in the specified row
    For idxCol = 1 To wsInt.Cells(RowTask5WithdrawBranchXAxis, wsInt.columns.Count).End(xlToLeft).Column
        ' Get the value of the current cell
        cellKey = wsInt.Cells(RowTask5WithdrawBranchXAxis, idxCol).Value
        cellValue = Trim(wsInt.Cells(RowTask5WithdrawBranchYAxis, idxCol).Value)
        If cellValue <> "" Then
            ' Process the cell value (replace With your code)
            ' Debug.Print "Column: " & idxCol & ", Value: " & cellValue
            hasMulti = True
            strHits = strHits & cellKey & ": " & cellValue & vbNewLine
        End If

    Next idxCol

    ' 寫摘要
    If hasMulti = True Then
        Range(posTask5Result).Value = "在當日臨櫃提領現金交易數據中，在觀察期間有發現不同地點提領: " & vbNewLine & strHits
    End If

    ' For task 7
    lastColumn = wsInt.Cells(RowTask7DiffXAxis, wsInt.columns.Count).End(xlToLeft).Column
    Dim cntDiffLess As Double
    Dim totalDeposit As Double
    Dim totalTransfer As Double
    Dim ratio       As Double

    cntDiffLess = 0
    totalDeposit = 0
    totalTransfer = 0

    For i = 1 To lastColumn
        totalDeposit = totalDeposit + CDbl(wsInt.Cells(RowTask7DiffDeposit, i).Value)
        totalTransfer = totalTransfer + CDbl(wsInt.Cells(RowTask7DiffTransfer, i).Value)

        ratio = CDbl(wsInt.Cells(RowTask7DiffRatio, i).Value)
        If (ratio > 0) And (ratio <= 0.03) Then
            cntDiffLess = cntDiffLess + 1
        End If
    Next i

    If totalDeposit <> 0 Then
        ratio = Abs((totalDeposit - totalDeposit) / totalDeposit)
    Else
        ratio = 0
    End If
    If (ratio > 0) And (ratio < 0.03) Then        ' FIXME:make 0.03 a constant
        Range(posTask7_1Result).Value = "在觀察區間內，總出入帳金額差距甚小"
    End If

    If cntDiffLess > 0 Then
        Range(posTask7_2Result).Value = "在觀察區間內，發現 " & CStr(cntDiffLess) & "  次出入帳金額差距甚小"
    End If

    ' Task 7
    CreateBarChart RowTask7DiffXAxis, RowTask7DiffRatio, Range("H12:X13"), "出入帳金額比例", "差距比例"

End Sub

Function ProcessGroupByDay(groupRange As Range, colName As String, idxOut As Long) As Boolean
    Dim intResult   As Long
    intResult = groupRange.row + groupRange.Rows.Count - 1

    Dim sum         As Long
    Dim cell        As Range

    ' Debug.Print "Summary: " & sum
    Dim wsInt       As Worksheet
    Set wsInt = ThisWorkbook.Worksheets(SheetNameIntermediate)

    ' for task 4 & 5
    Dim cntSuspicious As Long
    Dim cntBranch   As Long
    cntBranch = 0

    Dim dictBranch  As New Scripting.Dictionary
    Dim strBranch   As String
    Dim strKey      As String
    Dim result      As Variant

    strBranch = ""

    For Each cell In groupRange.columns(ColBranchID).Cells
        If (cell.Text <> SelfServiceID) And (cell.Text <> SelfServiceID_2) Then
            result = LookupBranchInfo(cell.Text)
            If IsError(result) Then
                Debug.Print "branch info Not found" & cell.Text
                strKey = CStr(cell.Text)

            Else
                strKey = CStr(result)
            End If

            If Not dictBranch.Exists(strKey) Then
                cntBranch = cntBranch + 1
                dictBranch.Add strKey, 1
                strBranch = strBranch + strKey + ", "
            End If
        End If
    Next cell        ' for each ColBranchID's Cells

    If Len(strBranch) > 1 Then
        strBranch = Left(strBranch, Len(strBranch) - 2)
    End If

    Dim cntSuspcious As Long
    cntSuspcious = 0
    For Each cell In groupRange.columns(ColAmountTransfer).Cells
        If IsNumeric(cell.Value) Then
            If cell.Value >= 400000 And cell.Value < 500000 Then        'FIXME: this should be from setting, 不確定lower bound
                cntSuspicious = cntSuspicious + 1
                ExcelUtils.HighlightRow SheetNameSimple, cell.row, ReasonCloseTo50w, ColorRed
            End If
        End If
    Next cell

    If cntSuspicious > 0 Then
        wsInt.Cells(RowTask4WithdrawOverCounterXAxis, idxOut).Value = colName
        wsInt.Cells(RowTask4WithdrawOverCounterYAxis, idxOut).Value = cntSuspicious

        wsInt.Cells(RowTask5WithdrawBranchXAxis, idxOut).Value = colName
        If cntBranch >= 2 Then
            wsInt.Cells(RowTask5WithdrawBranchYAxis, idxOut).Value = strBranch
        End If

        ProcessGroupByDay = True
    Else
        ProcessGroupByDay = False
    End If

End Function

Function ProcessGroupByDaySmallAmount(groupRange As Range, colName As String, idxOut As Long) As Boolean
    Dim intResult   As Long
    intResult = groupRange.row + groupRange.Rows.Count - 1

    Dim cell        As Range

    Dim wsInt       As Worksheet
    Set wsInt = ThisWorkbook.Worksheets(SheetNameIntermediate)

    ' for Task 12, 13
    Dim cntOneDollar, cntHundred, cntInterBank As Long
    cntOneDollar = 0
    cntHundred = 0
    cntInterBank = 0
    Dim amount      As Long
    For Each cell In groupRange.columns(ColTSType).Cells
        amount = groupRange.Cells(cell.row - groupRange.row + 1, ColAmountTransfer).Value
        'FIXME: 跨行存款怎麼看
        If (cell.Text = "03跨行轉帳") And (amount > 0) Then
            cntInterBank = cntInterBank + 1

        End If

        If cell.Text = "03跨行轉帳" Or cell.Text = "04本行轉帳" Then
            If amount = 1 Then
                cntOneDollar = cntOneDollar + 1
                ExcelUtils.HighlightRow SheetNameSimple, cell.row, ReasonSmallTSAmount, ColorRed
            ElseIf amount > 1 And amount < 100 Then
                cntHundred = cntHundred + 1
                ExcelUtils.HighlightRow SheetNameSimple, cell.row, ReasonSmallTSAmount, ColorPink
            End If

        End If
    Next cell        ' for each ColTSType's Cells

    If cntOneDollar > 0 Then
        wsInt.Cells(RowTask12SmallTransferXAxis, idxOut).Value = colName
        wsInt.Cells(RowTask12SmallTransferOneDollar, idxOut).Value = cntOneDollar

    End If

    If cntHundred > 0 Then
        wsInt.Cells(RowTask12SmallTransferXAxis, idxOut).Value = colName
        wsInt.Cells(RowTask12SmallTransferHundred, idxOut).Value = cntHundred
    End If

    If cntInterBank >= 5 Then
        wsInt.Cells(RowTask12SmallTransferXAxis, idxOut).Value = colName
        wsInt.Cells(RowTask12SmallTransferMany, idxOut).Value = cntInterBank

    End If

    If cntOneDollar > 0 Or cntHundred > 0 Or cntInterBank >= 5 Then
        ProcessGroupByDaySmallAmount = True
    Else
        ProcessGroupByDaySmallAmount = False
    End If

End Function

Function ProcessGroupByDayDiff(groupRange As Range, colName As String, idxOut As Long) As Boolean
    Dim cell        As Range
    Dim wsInt       As Worksheet
    Dim sumDeposit  As Double
    Dim sumTransfer As Double
    Dim amtDeposit  As Double
    Dim amtTransfer As Double

    Set wsInt = ThisWorkbook.Worksheets(SheetNameIntermediate)
    sumDeposit = 0
    sumTransfer = 0

    ' for Task 12, 13
    For Each cell In groupRange.columns(ColAmountTransfer).Cells
        amtTransfer = CDbl(cell.Value)
        amtDeposit = CDbl(groupRange.Cells(cell.row - groupRange.row + 1, ColAmountDeposit).Value)

        sumTransfer = sumTransfer + amtTransfer
        sumDeposit = sumDeposit + amtDeposit

    Next cell        ' for each ColAmountTransfer's Cells

    'FIXME: set format as percentage!
    wsInt.Cells(RowTask7DiffXAxis, idxOut).Value = colName
    wsInt.Cells(RowTask7DiffDeposit, idxOut).Value = sumDeposit
    wsInt.Cells(RowTask7DiffTransfer, idxOut).Value = sumTransfer
    If sumDeposit <> 0 Then
        wsInt.Cells(RowTask7DiffRatio, idxOut).Value = Abs((sumDeposit - sumTransfer) / sumDeposit)
    Else
        wsInt.Cells(RowTask7DiffRatio, idxOut).Value = 0

    End If
    ProcessGroupByDayDiff = True

End Function

Sub GetVisitedBranch()
    Dim ws          As Worksheet
    Set ws = ThisWorkbook.Worksheets(SheetNameInputData)
    Dim wsInt       As Worksheet
    Set wsInt = ThisWorkbook.Worksheets(SheetNameIntermediate)
    Dim dictBranch  As New Scripting.Dictionary

    Dim lastRow     As Long
    Dim row         As Long
    Dim cell        As Range

    lastRow = ws.Cells(ws.Rows.Count, ColTSDate).End(xlUp).row
    ' Loop through rows starting from row rowDataBegin, day by day
    For row = RowDataBegin To lastRow        ' lastRow FIXME: For debug only
        Set cell = ws.Cells(row, ColBranchID)

        If (cell.Text <> SelfServiceID) And (cell.Text <> SelfServiceID_2) Then
            If Not dictBranch.Exists(cell.Text) Then
                dictBranch.Add cell.Text, 1
            Else
                dictBranch(cell.Text) = dictBranch(cell.Text) + 1
            End If
        End If

    Next row

    ' Convert the dictionary To arrays
    Dim result      As Variant
    Dim strKey      As String

    Dim keysArray() As Variant
    Dim valuesArray() As Variant
    Dim key         As Variant

    ReDim keysArray(0)
    ReDim valuesArray(0)
    If dictBranch.Count > 0 Then
        ReDim keysArray(1 To dictBranch.Count)
        ReDim valuesArray(1 To dictBranch.Count)
        Dim i           As Long
        i = 1
        For Each key In dictBranch.keys
            keysArray(i) = key
            valuesArray(i) = dictBranch(key)
            i = i + 1
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

            wsInt.Cells(RowTask5VisitedBranchXAxis, i).Value = "分行代碼:" & Format(strKey, "0000")
            wsInt.Cells(RowTask5VisitedBranchYAxis, i).Value = valuesArray(i)

        Next i
    End If

    CreateBarChart RowTask5VisitedBranchXAxis, RowTask5VisitedBranchYAxis, Range("Q7:X8"), "去過的分行統計", "所有交易次數"

End Sub

Sub GetTransferAmount()
    Dim ws          As Worksheet
    Set ws = ThisWorkbook.Worksheets(SheetNameInputData)
    Dim wsInt       As Worksheet
    Set wsInt = ThisWorkbook.Worksheets(SheetNameIntermediate)
    Dim dictTransfer  As New Scripting.Dictionary
    Dim dictTransferCount  As New Scripting.Dictionary
    Dim dictDeposit  As New Scripting.Dictionary

    Dim lastRow     As Long
    Dim row         As Long
    Dim cell        As Range
    Dim val         As Double

    lastRow = ws.Cells(ws.Rows.Count, ColTSDate).End(xlUp).row
    ' Loop through rows starting from row rowDataBegin, day by day
    For row = RowDataBegin To lastRow
        Set cell = ws.Cells(row, colAccount)

        If Not IsEmpty(cell) Then
            val = CDbl(ws.Cells(row, ColAmountTransfer).Value)
            If val > 0 Then
                If Not dictTransfer.Exists(cell.Text) Then
                    dictTransfer.Add cell.Text, val
                    dictTransferCount.Add cell.Text, 1
                Else
                    dictTransfer(cell.Text) = dictTransfer(cell.Text) + val
                    dictTransferCount(cell.Text) = dictTransferCount(cell.Text) + 1
                End If
            End If

            val = CDbl(ws.Cells(row, ColAmountDeposit).Value)
            If val > 0 Then
                If Not dictDeposit.Exists(cell.Text) Then
                    dictDeposit.Add cell.Text, val
                Else
                    dictDeposit(cell.Text) = dictDeposit(cell.Text) + val
                End If
            End If
        End If

    Next row

    If dictTransfer.Count <= 0 Then
        GoTo EndOfSub
    End If

    ' Convert the dictionary To arrays
    Dim result      As Variant
    Dim strKey      As String

    Dim keysArray() As Variant
    Dim valuesArray() As Variant
    Dim key         As Variant
    Dim i           As Long

    ReDim keysArray(1 To dictTransfer.Count)
    ReDim valuesArray(1 To dictTransfer.Count)
    i = 1
    For Each key In dictTransfer.keys
        keysArray(i) = key
        valuesArray(i) = dictTransfer(key)
        i = i + 1
    Next key

    ' Sort the arrays by values
    Utils.QuickSort valuesArray, keysArray, LBound(valuesArray), UBound(valuesArray)

    ' Write the sorted arrays
    For i = LBound(valuesArray) To UBound(valuesArray)

        With wsInt.Cells(RowTopCounterparty + i - 1, 1)
            .NumberFormat = "@"
            .Font.name = "微軟正黑體"
            .Font.Size = "11"
            .Value = keysArray(i)
        End With
        With wsInt.Cells(RowTopCounterparty + i - 1, 2)
            .NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* "" - ""??_);_(@_)"
            .Font.name = "Arial"
            .Font.Size = "11"
            .HorizontalAlignment = xlRight

            .Value = valuesArray(i)
        End With
        With wsInt.Cells(RowTopCounterparty + i - 1, 3)
            .NumberFormat = "_(* #,##0_);_(* (#,##0);_(* "" - ""??_);_(@_)"
            .Font.name = "Arial"
            .Font.Size = "11"
            .HorizontalAlignment = xlRight

            .Value = dictTransferCount(keysArray(i))
        End With
    Next i

    ReDim keysArray(1 To dictDeposit.Count)
    ReDim valuesArray(1 To dictDeposit.Count)
    i = 1
    For Each key In dictDeposit.keys
        keysArray(i) = key
        valuesArray(i) = dictDeposit(key)
        i = i + 1
    Next key

    ' Sort the arrays by values
    Utils.QuickSort valuesArray, keysArray, LBound(valuesArray), UBound(valuesArray)

    ' Write the sorted arrays
    For i = LBound(valuesArray) To UBound(valuesArray)

        With wsInt.Cells(RowTopCounterparty + i - 1, 4)
            .NumberFormat = "@"
            .Font.name = "微軟正黑體"
            .Font.Size = "11"
            .Value = keysArray(i)
        End With
        With wsInt.Cells(RowTopCounterparty + i - 1, 5)
            .NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* "" - ""??_);_(@_)"
            .Font.name = "Arial"
            .Font.Size = "11"
            .HorizontalAlignment = xlRight

            .Value = valuesArray(i)
        End With

    Next i
EndOfSub:
    'FIXME: 區分跨行
End Sub

'/* 如何判斷是ATM交易:  1. 看櫃員代碼 (慢) 2. 收款機/ATM 欄 */
' or cash
' TASK ATM 存現 地緣性 by day
Sub ProcessRowsATM()
    Dim ws          As Worksheet
    Set ws = ThisWorkbook.Worksheets(SheetNameInputData)
    Dim wsInt       As Worksheet
    Set wsInt = ThisWorkbook.Worksheets(SheetNameIntermediate)

    Dim lastRow     As Long
    Dim currentGroup As Range
    Dim row         As Long
    Dim idxcolOutput As Integer
    Dim currentYear, currentMonth, currentDay As Integer
    Dim idxcolQuick As Integer

    idxcolOutput = 1        ' output column index
    idxcolQuick = 1
    lastRow = ws.Cells(ws.Rows.Count, ColTSDate).End(xlUp).row
    ' Loop through rows starting from row rowDataBegin, day by day
    For row = RowDataBegin To lastRow
        Dim dateValue As Date
        Dim y, m, d As Integer
        Dim ret     As Boolean

        dateValue = ws.Cells(row, ColTSDate).Value
        y = year(dateValue)
        m = month(dateValue)
        d = Day(dateValue)
        ' Check If the date has changed Or If it's the first row
        If row = RowDataBegin Or y <> currentYear Or m <> currentMonth Or d <> currentDay Then
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
            Set currentGroup = ws.Range(ws.Cells(row, 1), ws.Cells(row, MaxSrcCol))

            currentYear = y
            currentMonth = m
            currentDay = d
        Else
            ' Add the row To the current group
            Set currentGroup = Union(currentGroup, ws.Range(ws.Cells(row, 1), ws.Cells(row, MaxSrcCol)))

        End If
    Next row
    ' Process the last group of rows
    If Not currentGroup Is Nothing Then
        ProcessGroupByATM currentGroup, CStr(currentYear) & "/" & CStr(currentMonth) & "/" & CStr(currentDay), idxcolOutput
        ProcessGroupByCashQuickInOut currentGroup, CStr(currentYear) & "/" & CStr(currentMonth) & "/" & CStr(currentDay), idxcolQuick
    End If

    ' 寫摘要
    gCntMultiATM = 0
    gCntMultiDeposit = 0
    Dim lastCol     As Long
    lastCol = wsInt.Cells(RowTask9ATMDepositXAxis, wsInt.columns.Count).End(xlToLeft).Column

    Dim i           As Long
    For i = 1 To lastCol
        If wsInt.Cells(RowTask9ATMDepositYAxis, i) >= 10 Then        ' FIXME:constants
            Utils.CountOne gCntMultiDeposit
        End If

        If wsInt.Cells(RowTask9ATMDepositLocation, i) >= 3 Then        ' FIXME:constants
            Utils.CountOne gCntMultiATM
        End If
    Next i

    If gCntMultiDeposit > 0 Then
        Range(posTask10Result).WrapText = True
        Range(posTask10Result).Value = "在觀察區間發生了 " & CStr(gCntMultiDeposit) & " 次單日十次以上ATM 存現案例。"
    End If

    If gCntMultiATM > 0 Then
        Range(posTask9Result).WrapText = True
        Range(posTask9Result).Value = "在觀察區間發生了 " & CStr(gCntMultiATM) & " 次單日多個 ATM 存現案例"
    End If

    CreateBarChart RowTask9ATMDepositXAxis, RowTask9ATMDepositYAxis, Range("H15:X16"), "多個ＡＴＭ存現", "次數"

    lastCol = wsInt.Cells(RowTask6QuickXAxis, wsInt.columns.Count).End(xlToLeft).Column

    Range(posTask6Result).WrapText = True
    Range(posTask6Result).Value = "在觀察區間發生了 " & CStr(lastCol) & " 次快進快出案例"

    CreateBarChart RowTask6QuickXAxis, RowTask6QuickYAxis, Range("H10:X10"), "快進快出", "次數"

End Sub

Function ProcessGroupByATM(groupRange As Range, colName As String, idxOut As Integer) As Boolean
    Dim cell        As Range

    Dim wsInt       As Worksheet
    Set wsInt = ThisWorkbook.Worksheets(SheetNameIntermediate)

    Dim dictATM     As New Scripting.Dictionary
    Dim cntATM      As Integer

    Dim strChannel  As String
    Dim strBankID   As String
    Dim strBranch   As String
    Dim cntATMTrans As Long
    strBranch = ""
    cntATMTrans = 0
    Dim amount      As Double

    For Each cell In groupRange.columns(ColBranchID).Cells
        If (cell.Text = SelfServiceID) Or (cell.Text = SelfServiceID_2) Then
            strChannel = groupRange.Cells(cell.row - groupRange.row + 1, colChannel).Text
            amount = CDbl(groupRange.Cells(cell.row - groupRange.row + 1, ColAmountDeposit).Value)

            If amount > 0 Then
                If strChannel = "存款機" Or strChannel = "ＡＴＭ" Or strChannel = "ATM" Then

                    cntATMTrans = cntATMTrans + 1

                    If Not dictATM.Exists(strChannel) Then
                        dictATM.Add strChannel, 1
                    Else
                        dictATM(strChannel) = dictATM(strChannel) + 1
                    End If
                ElseIf Len(strChannel) < 4 And Utils.ContainsNumbers(strChannel) = True Then
                    strBankID = Format(CInt(strChannel), "000")

                    cntATMTrans = cntATMTrans + 1

                    If Not dictATM.Exists(strBankID) Then
                        dictATM.Add strBankID, 1
                    Else
                        dictATM(strBankID) = dictATM(strBankID) + 1
                    End If
                End If
            End If
        End If        ' Check deposit amount > 0 only
ContinueLoop:
    Next cell
    Dim result      As String
    If (dictATM.Count >= 3) Or (cntATMTrans >= 10) Then        ' FIXME: no hard-coding values
        result = Join(dictATM.keys, ";")
        wsInt.Cells(RowTask9ATMDepositXAxis, idxOut).Value = colName
        wsInt.Cells(RowTask9ATMDepositYAxis, idxOut).Value = cntATMTrans
        wsInt.Cells(RowTask9ATMDepositLocation, idxOut).Value = dictATM.Count
        wsInt.Cells(RowTask9ATMDepositLocationDetail, idxOut).Value = result
        'Debug.Print "Write " & colName & " " & CStr(cntATMTrans) & " " & CStr(dictATM.Count) & " " & result

        ProcessGroupByATM = True
    Else
        ProcessGroupByATM = False
    End If

End Function

Function ProcessGroupByCashQuickInOut(groupRange As Range, colName As String, idxOut As Integer) As Boolean
    Dim cell        As Range

    Dim wsInt       As Worksheet
    Set wsInt = ThisWorkbook.Worksheets(SheetNameIntermediate)

    Dim currentTime As Date
    Dim bInit       As Boolean
    Dim t           As Date
    Dim tDiff       As Double
    Dim bQuick      As Boolean
    Dim cntQuick    As Integer
    Dim amtTransfer     As Double
    Dim amtDeposit      As Double
    Dim bTransfer   As Boolean
    Dim bDeposit    As Boolean

    bInit = False
    bQuick = False
    cntQuick = 0

    For Each cell In groupRange.columns(ColTSType).Cells
        If cell.Text = "01現金交易" Then
            amtTransfer = CDbl(groupRange.Cells(cell.row - groupRange.row + 1, ColAmountTransfer).Value)
            If amtTransfer > 0 Then
                bTransfer = True
            End If
            amtDeposit = CDbl(groupRange.Cells(cell.row - groupRange.row + 1, ColAmountDeposit).Value)
            If amtDeposit > 0 Then
                bDeposit = True
            End If
            If bInit = True Then
                t = groupRange.Cells(cell.row - groupRange.row + 1, ColTSTime).Value
                tDiff = DateDiff("n", currentTime, t)
                If t <= 10 Then        'FIXME: 10 minutes constant
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
                currentTime = groupRange.Cells(cell.row - groupRange.row + 1, ColTSTime).Value
                bInit = True
            End If
            'Debug.Print "amount " & amount
        End If

    Next cell

    If (cntQuick > 1) And (bTransfer = True) And (bDeposit = True) Then
        wsInt.Cells(RowTask6QuickXAxis, idxOut).Value = colName
        wsInt.Cells(RowTask6QuickYAxis, idxOut).Value = cntQuick
        ProcessGroupByCashQuickInOut = True
    Else
        ProcessGroupByCashQuickInOut = False
    End If

End Function

Sub ProcessRowsBalance()
    Dim ws          As Worksheet
    Dim lastRow     As Long
    Dim currentDate As Date
    Dim currentBalance As Long
    Dim targetCol   As Long
    Dim i           As Long

    Set ws = ThisWorkbook.Sheets(SheetNameInputData)

    lastRow = ws.Cells(ws.Rows.Count, ColTSDate).End(xlUp).row

    Dim wsInt       As Worksheet
    Set wsInt = ThisWorkbook.Worksheets(SheetNameIntermediate)

    ' Initialize variables
    currentDate = ws.Cells(RowDataBegin, ColTSDate).Value        ' Assuming the first date is in A9
    If VarType(ws.Cells(RowDataBegin, colBalance).Value) <> vbString Then
        currentBalance = ws.Cells(RowDataBegin, colBalance).Value        ' Assuming the first balance is in I9
    Else
        currentBalance = 0
    End If
    targetCol = 1
    ' Loop through each row starting from the second row
    For i = RowDataBegin To lastRow
        If VarType(ws.Cells(i, colBalance).Value) <> vbString Then        ' workaround for data error (欄位飄掉)
            ' Check if the date is the same
            If ws.Cells(i, ColTSDate).Value = currentDate Then
                ' Update the current balance
                currentBalance = CDbl(ws.Cells(i, colBalance).Value)
            Else
                ' Write the date and balance to the new rows
                wsInt.Cells(RowTask11BalanceXAxis, targetCol).Value = currentDate
                wsInt.Cells(RowTask11BalanceYAxis, targetCol).Value = currentBalance

                ' Update currentDate and currentBalance for the new date
                currentDate = ws.Cells(i, ColTSDate).Value
                currentBalance = CDbl(ws.Cells(i, colBalance).Value)
                ' Move to the next row
                targetCol = targetCol + 1
            End If
        End If
    Next i
    ' Write the last date and balance to the final row

    wsInt.Cells(RowTask11BalanceXAxis, targetCol).Value = currentDate
    wsInt.Cells(RowTask11BalanceYAxis, targetCol).Value = currentBalance

    CreateBarChart RowTask11BalanceXAxis, RowTask11BalanceYAxis, Range("H17:X17"), "日期", "餘額"

    ' Compute the date difference

    Dim startDate       As Date
    Dim endDate         As Date
    Dim dayDiff, dayTotal As Long
    Dim idxDate         As Integer
    Dim currentValue    As Long
    Dim blUnderThreshold As Boolean
    Dim cntUnderDays    As Integer

    Dim lastCol         As Integer

    lastCol = wsInt.Cells(RowTask11BalanceXAxis, ws.columns.Count).End(xlToLeft).Column

    startDate = wsInt.Cells(RowTask11BalanceXAxis, 1).Value
    endDate = wsInt.Cells(RowTask11BalanceXAxis, lastCol).Value
    dayDiff = DateDiff("d", startDate, endDate)

    dayTotal = dayDiff + 1
    idxDate = 1

    cntUnderDays = 0
    currentDate = wsInt.Cells(RowTask11BalanceXAxis, idxDate).Value
    currentValue = wsInt.Cells(RowTask11BalanceYAxis, idxDate).Value
    If currentValue <= 1000 Then        ' FIXME: make it constant]
        blUnderThreshold = True
    Else
        blUnderThreshold = False
    End If
    Dim d               As Date
    For i = 1 To dayTotal
        d = DateAdd("d", (i - 1), startDate)
        If blUnderThreshold = True Then
            cntUnderDays = cntUnderDays + 1
        End If

        If d = currentDate Then
            ' matched

            idxDate = idxDate + 1
            currentDate = wsInt.Cells(RowTask11BalanceXAxis, idxDate).Value
            currentValue = wsInt.Cells(RowTask11BalanceYAxis, idxDate).Value
            If currentValue <= 1000 Then        ' FIXME: make it constant
                blUnderThreshold = True
            Else
                blUnderThreshold = False
            End If

        End If

    Next i

    Range(posTask11Result).WrapText = True
    Range(posTask11Result).Value = "在觀察區間 " & CStr(dayTotal) & " 天中，有 " & CStr(cntUnderDays) & " 天每日存款餘額少於1000元。"
    Debug.Print "total days Is " & CStr(dayTotal) & " under days Is " & CStr(cntUnderDays)

End Sub

Private Sub NewAnalyze_Click()

    Application.ScreenUpdating = False        ' for performance
    UpdateSimplePage = False
    ' 準備資料
    FormatData

    '分析資料
    Analyze

    '準備金流及交易對手
    Application.DisplayAlerts = False

    CreatePivotTables
    Application.DisplayAlerts = True

    Application.ScreenUpdating = True        ' for performance

    ThisWorkbook.Sheets(SheetNameSimple).Activate
    'ThisWorkbook.Sheets("3.2金流與交易對手").Activate

    ActiveWindow.ScrollRow = 1

End Sub

Private Sub Analyze()
    Dim ws          As Worksheet
    Dim wsSimple    As Worksheet
    Dim wsInput     As Worksheet

    ExtraData.InitDicts

    Set ws = ThisWorkbook.Sheets(SheetNameMain)
    Set wsSimple = ThisWorkbook.Sheets(SheetNameSimple)
    Set wsInput = ThisWorkbook.Sheets(SheetNameInputData)

    ws.Range(posRunStatus).Value = "清除前次分析資料"

    ExcelUtils.RemoveAllCharts SheetNameMain

    ExcelUtils.ClearSheet SheetNameIntermediate

    ExcelUtils.ClearSheet SheetNameSimple, False

    ExcelUtils.ClearSheet SheetNameMoney

    ws.Range("F2:F30").ClearContents

    ws.Range(posRunStatus).Value = "準備交易簡化版資料"
    ' Assume source data is ready
    Dim arrDup()    As Variant
    arrDup = Array(1, 16, 4, 17, 18, 9, 10, 14, 13, 27)

    FillSimplePage arrDup

    GetTransferAmount

    ' Fill customer info
    'wsSimple.Range("A2").Value = wsInput.Range("A4").Value
    'wsSimple.Range("A3").Value = wsInput.Range("B4").Value
    ' wsSimple.Range("A5").Value = wsInput.Range("A5").Value
    'wsSimple.Range("A6").NumberFormat = "@"
    ' wsSimple.Range("A6").Value = Application.text(wsInput.Range("B5").Value, "0")

    ws.Activate
    ws.Range(posRunStatus).Value = "開始分析"

    ProcessRowsByMonth
    ProcessRowsByDay
    GetVisitedBranch
    ProcessRowsATM
    ProcessRowsBalance
    ws.Range(posRunStatus).Value = "分析完成"
    Debug.Print "End of Program"

End Sub

' prepare data in the simple page, put all UI related here
Sub FillSimplePage(arrColumnIDs() As Variant)
    Dim wsSource    As Worksheet
    Dim wsDestination As Worksheet
    Dim colID       As Variant
    Dim destinationCol As Long
    Dim lastCol     As String
    Dim timeCol, warnCol As String
    ' Set the source and destination worksheets
    Set wsSource = ThisWorkbook.Sheets(SheetNameInputData)
    Set wsDestination = ThisWorkbook.Sheets(SheetNameSimple)

    lastCol = "M"        ' Attention: Check here when columns are changed
    timeCol = "C"
    warnCol = "K"

    With wsDestination.columns("A")
        .NumberFormat = "yyyy/mm/dd"
        .Font.name = "微軟正黑體"
        .Font.Size = "11"
        .HorizontalAlignment = xlLeft
    End With

    With wsDestination.columns("B")
        .NumberFormat = "@"
        .Font.name = "微軟正黑體"
        .Font.Size = "11"
    End With

    With wsDestination.columns(timeCol)
        .NumberFormat = "hh:mm:ss"
        .HorizontalAlignment = xlCenter
        .Font.name = "微軟正黑體"
        .Font.Size = "11"
    End With

    With wsDestination.columns("D:F")
        .NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* "" - ""??_);_(@_)"
        .Font.name = "Arial"
        .Font.Size = "11"
        .HorizontalAlignment = xlRight
    End With

    With wsDestination.columns("G:M")
        .NumberFormat = "@"
        .Font.name = "微軟正黑體"
        .Font.Size = "11"
    End With

    ' Start duplicating columns from the 7th column in the destination sheet
    destinationCol = 1
    Dim i           As Long
    Dim startRow, lastRow As Long

    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).row
    startRow = RowDataBegin
    ' Copy source columns
    For Each colID In arrColumnIDs
        If colID >= 1 And colID <= wsSource.columns.Count Then
            wsDestination.Range(wsDestination.Cells(RowDataBegin - 1, destinationCol), wsDestination.Cells(lastRow, destinationCol)).Value = _
                                                                 wsSource.Range(wsSource.Cells(RowDataBegin - 1, colID), wsSource.Cells(lastRow, colID)).Value

            wsDestination.columns(destinationCol).AutoFit
            destinationCol = destinationCol + 1
        End If

    Next colID

    ' used in checking virtual accounts and channels
    wsDestination.Range(wsDestination.Cells(RowDataBegin - 1, "Y"), wsDestination.Cells(lastRow, "Y")).Value = _
    wsSource.Range(wsSource.Cells(RowDataBegin - 1, "AB"), wsSource.Cells(lastRow, "AB")).Value
    wsDestination.Range(wsDestination.Cells(RowDataBegin - 1, "Z"), wsDestination.Cells(lastRow, "Z")).Value = _
                                                         wsSource.Range(wsSource.Cells(RowDataBegin - 1, "S"), wsSource.Cells(lastRow, "S")).Value

    ' set header
    wsDestination.Range(warnCol & "8").Value = "顏色標註理由"
    wsDestination.columns(warnCol).AutoFit
    wsDestination.Range("L8").Value = "表徵"
    wsDestination.Range("M8").Value = "調查結果"

    Dim cellTime    As Date

    ' Check and highlight
    Dim strAcc      As String
    For i = startRow To lastRow
        ' set background color
        If (i - startRow) Mod 2 = 0 Then
            wsDestination.Range(wsDestination.Cells(i, "A"), wsDestination.Cells(i, "K")).Interior.color = ColorYellow
        End If
        ' Color transations in the morning FIXME: move to viewer
        If VarType(wsDestination.Cells(i, timeCol).Value) = vbDouble Then
            cellTime = wsDestination.Cells(i, timeCol).Value
            If cellTime >= TimeValue("0:00:00") And cellTime <= TimeValue("4:00:00") Then
                ExcelUtils.HighlightRow SheetNameSimple, i, ReasonTSInMorning, ColorBlue
            End If
        End If
        ' Check amount
        If (Not IsEmpty(wsDestination.Cells(i, "D")) Or Not IsEmpty(wsDestination.Cells(i, "E"))) Then
            If IsNumeric(wsDestination.Cells(i, "D").Value) And wsDestination.Cells(i, "D").Value > 100000 Then
                If wsDestination.Cells(i, "D").Value > 500000 Then
                    ExcelUtils.HighlightRow SheetNameSimple, i, ReasonTSLargeAmount, ColorPink
                Else
                    ExcelUtils.HighlightRow SheetNameSimple, i, ReasonTSLargeAmount, ColorOrange
                End If

            ElseIf IsNumeric(wsDestination.Cells(i, "E").Value) And wsDestination.Cells(i, "E").Value > 100000 Then
                If wsDestination.Cells(i, "E").Value > 500000 Then
                    ExcelUtils.HighlightRow SheetNameSimple, i, ReasonTSLargeAmount, ColorPink
                Else
                    ExcelUtils.HighlightRow SheetNameSimple, i, ReasonTSLargeAmount, ColorOrange
                End If
            End If
        End If

        ' check blacklist
        strAcc = wsDestination.Cells(i, "G")
        If Not IsEmpty(strAcc) And VarType(strAcc) = vbString Then
            If DictBlacklist.Exists(strAcc) Then
                'Debug.Print " 警示戶　hit!!! " & strAcc
                ExcelUtils.HighlightRow SheetNameSimple, i, ReasonDWLAlert, ColorRed
            End If
        End If

    Next i

    ' FIXME: pass object
    MergeCellsByDate

    ' Disable cell warning
    Application.ErrorCheckingOptions.NumberAsText = False

    ' Set header style
    With wsDestination.Range("A8:" & lastCol & "8")
        .Font.color = ColorWhite
        .Font.Bold = True
        .Interior.color = ColorGreen        ' Header color
        .HorizontalAlignment = xlCenter
    End With

    wsDestination.Range("A8:" & lastCol & CStr(lastRow)).AutoFilter

    'freeze
    wsDestination.Activate
    With ActiveWindow
        If .FreezePanes Then .FreezePanes = False
        .SplitRow = 8
        .FreezePanes = True
    End With

End Sub

Sub SetEditboxText(ctrlName As String, Text As String)

    Dim ws          As Worksheet
    Dim edit        As OLEObject

    Set ws = ThisWorkbook.Sheets(SheetNameSimple)

    Set edit = ws.OLEObjects(ctrlName)

    edit.Object.Text = Text

End Sub

Sub MergeCellsByDate()
    Dim lastRow     As Long
    Dim currentCell As Range
    Dim currentDate As String
    Dim wsDestination As Worksheet
    Set wsDestination = ThisWorkbook.Sheets(SheetNameSimple)
    Dim color       As Long
    color = 1

    ' Find the last row with data in column A
    lastRow = wsDestination.Cells(Rows.Count, "A").End(xlUp).row
    ' Start from row 9 since the real data starts from there

    Application.DisplayAlerts = False
    For Each currentCell In wsDestination.Range("A9:A" & lastRow)

        'Debug.Print "current: " & currentDate & " now: " & currentCell.Value
        If Trim(currentCell.Value) = currentDate And Not IsEmpty(currentCell.Value) Then
            ' Merge the cells without warning
            wsDestination.Range(currentCell, currentCell.Offset(-1, 0)).Merge
            wsDestination.Range(currentCell, currentCell.Offset(-1, 0)).VerticalAlignment = xlCenter

            If (color) Mod 2 = 0 Then
                wsDestination.Range(currentCell, currentCell.Offset(-1, 0)).Interior.color = ColorYellow
            Else
                wsDestination.Range(currentCell, currentCell.Offset(-1, 0)).Interior.color = ColorWhite
            End If
        Else
            ' Update the current date
            currentDate = Trim(currentCell.Value)
            color = color + 1
            If (color) Mod 2 = 0 Then
                currentCell.Interior.color = ColorYellow
            Else
                currentCell.Interior.color = ColorWhite
            End If

        End If

    Next currentCell
    Application.DisplayAlerts = True
End Sub

Private Sub FormatData()
    Dim wsIn        As Worksheet
    Dim wsOut       As Worksheet

    ExtraData.InitDicts

    Set wsIn = ThisWorkbook.Sheets(SheetNameOrginal)
    Set wsOut = ThisWorkbook.Sheets(SheetNameInputData)

    Range(posRunStatus).Value = "開始整理"

    ExcelUtils.ClearSheet SheetNameInputData

    If wsIn.Range("A1") = "" And InStr(1, wsIn.Range("A4"), "客戶") = 1 Then
        MainAccName = wsIn.Range("B4")
    End If

    If wsIn.Range("A1") = "" And InStr(1, wsIn.Range("A5"), "帳號") = 1 Then
        MainAccId = wsIn.Range("B5")
    End If

    If wsIn.Range("A1") = "交易日期" Then
        wsIn.Rows("1:7").Insert Shift:=xlDown
    End If

    Dim i           As Long
    Dim lastRow     As Long

    Dim arrSrc()    As Variant
    Dim arrDst()    As Variant

    arrSrc = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "O", "P")        ' NOTE: N 欄沒有map
    arrDst = Array("A", "B", "C", "D", "E", "F", "G", "Q", "R", "I", "J", "K", "L", "N", "M")

    lastRow = wsIn.Cells(wsIn.Rows.Count, 1).End(xlUp).row

    With wsOut.columns("A:B")
        .NumberFormat = "yyyy/mm/dd"
        .Font.name = "微軟正黑體"
        .Font.Size = "11"
    End With

    With wsOut.columns("C:G")
        .NumberFormat = "@"
        .Font.name = "微軟正黑體"
        .Font.Size = "11"
    End With

    With wsOut.columns("H:I")
        .NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* "" - ""??_);_(@_)"
        .Font.name = "Arial"
        .Font.Size = "11"
        .HorizontalAlignment = xlRight
    End With

    With wsOut.columns("J:O")
        .NumberFormat = "@"
        .Font.name = "微軟正黑體"
        .Font.Size = "11"
    End With

    With wsOut.columns("P")
        .NumberFormat = "General"
        .Font.name = "微軟正黑體"
        .Font.Size = "11"
    End With

    With wsOut.columns("Q:R")
        .NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* "" - ""??_);_(@_)"
        .Font.name = "Arial"
        .Font.Size = "11"
        .HorizontalAlignment = xlRight
    End With

    With wsOut.columns("S:AC")
        .NumberFormat = "@"
        .Font.name = "微軟正黑體"
        .Font.Size = "11"
        .HorizontalAlignment = xlLeft
    End With

    ' Loop through each column ID in the array
    For i = 0 To UBound(arrSrc)
        wsOut.Range(wsOut.Range(arrDst(i) & "1"), wsOut.Range(arrDst(i) & CStr(lastRow))).Value = _
                                          wsIn.Range(wsIn.Range(arrSrc(i) & "1"), wsIn.Range(arrSrc(i) & CStr(lastRow))).Value

        wsOut.columns(arrDst(i)).AutoFit
    Next i

    ' Prepare headers
    wsOut.Cells(8, "H").Value = "Amt"
    wsOut.Cells(8, "O").Value = "交易月份"
    wsOut.Cells(8, "P").Value = "交易摘要"
    wsOut.Cells(8, "S").Value = "銀行代碼"
    wsOut.Cells(8, "T").Value = "TranType"
    wsOut.Cells(8, "U").Value = "ATM地點"
    wsOut.Cells(8, "V").Value = "ATM縣市"
    wsOut.Cells(8, "W").Value = "ATM區域"
    wsOut.Cells(8, "X").Value = "分行名"
    wsOut.Cells(8, "Y").Value = "分行縣市"
    wsOut.Cells(8, "Z").Value = "分行區域"
    wsOut.Cells(8, "AA").Value = "交易地點"
    wsOut.Cells(8, "AB").Value = "交易通路"
    wsOut.Cells(8, "AC").Value = "簡化交易時間"

    Dim wsATM          As Worksheet
    Dim dictATMCity As Object
    Dim dictATMLoc  As Object
    Dim wsBranch    As Worksheet
    Dim dictBranchName As Object
    Dim dictBranchCity As Object

    Set wsATM = ThisWorkbook.Sheets("ATM清單")
    Set dictATMCity = ExtraData.GetKVPairsToDict(wsATM, "C", "M")
    Set dictATMLoc = ExtraData.GetKVPairsToDict(wsATM, "C", "E")

    Set wsBranch = ThisWorkbook.Sheets("分行清單")
    Set dictBranchName = ExtraData.GetKVPairsToDict(wsBranch, "A", "D")
    Set dictBranchCity = ExtraData.GetKVPairsToDict(wsBranch, "A", "F")

    Dim tmpstr, strTmp As String
    Dim amtWithdraw As Long
    Dim amtDeposit  As Long
    Dim strSummary  As String
    Dim strSerial   As String
    Dim strNote     As String
    Dim strStore    As String
    Dim strClerk    As String
    Dim strTrans    As String
    Dim strAccount  As String

    For i = lastRow To 9 Step -1        ' FIXBUG: hard coding value
        strStore = wsIn.Cells(i, "E").Value
        strClerk = wsIn.Cells(i, "F").Value
        strSummary = wsIn.Cells(i, "G").Value

        amtWithdraw = wsIn.Cells(i, "H").Value
        amtDeposit = wsIn.Cells(i, "I").Value
        strAccount = wsIn.Cells(i, "K")
        strSerial = wsIn.Cells(i, "M").Value
        strNote = wsIn.Cells(i, "P").Value

        wsOut.Cells(i, "O").Value = Left(wsIn.Cells(i, "A").Text, 7)
        wsOut.Cells(i, "AC").Value = Left(wsIn.Cells(i, "D").Text, 2) & "時"

        If strSummary = "手續費" Or strSummary = "利息" Then
            wsOut.Rows(i).EntireRow.Delete
            GoTo GoSkip
        End If

        If amtWithdraw <> 0 Then
            wsOut.Cells(i, "H").Value = amtWithdraw
        ElseIf amtDeposit <> 0 Then
            wsOut.Cells(i, "H").Value = amtDeposit
        End If

        'GoTo GoSkip
        '"交易摘要"欄位
        strTrans = ""
        If Len(strSummary) > 0 Then
            If Left(strSummary, 1) = "跨" Then
                strTrans = "03跨行轉帳"
            ElseIf Right(strSummary, 1) = "匯" Then
                strTrans = "02匯款交易"
            ElseIf Left(strSummary, 1) = "現" Then
                strTrans = "01現金交易"
            ElseIf Left(strSummary, 1) = "轉" Then
                strTrans = "04本行轉帳"
            Else
                strTrans = "其他交易"
            End If
        End If
        wsOut.Cells(i, "P").Value = strTrans

        '"銀行代碼"欄位
        If strSerial = "" Then
            wsOut.Cells(i, "S").Value = ""
        Else
            wsOut.Cells(i, "S").Value = Left(strSerial, 3)
            ' If strAccount <> "" Then
            '     wsOut.Cells(i, "S").Value = wsOut.Cells(i, "S").Value & "-" & strAccount
            '  End If
        End If

        ' 交易通路
        Dim strChan As String

        strChan = ""
        If strNote <> "" Then
            If strNote = "行動網" Then
                strChan = Trim("行動網" & " " & strChan)
            ElseIf strNote = "網銀" Then
                strChan = Trim("網銀" & " " & strChan)
            ElseIf strNote = "ATM" Or strNote = "存款機" Or strNote = "ＡＴＭ" Then
                strChan = Trim("ATM自動化設備" & " " & strChan)
            ElseIf strNote = "收付網" Then
                strChan = Trim("收付網" & " " & strChan)
            ElseIf strNote = "證券款" Then
                strChan = Trim("證券款" & " " & strChan)
            ElseIf strNote = "傳真銀" Then
                strChan = Trim("傳真銀" & " " & strChan)
            ElseIf strNote = "ＸＭＬ" Then
                strChan = Trim("ＸＭＬ" & " " & strChan)
            ElseIf strNote = "FEDI" Then
                strChan = Trim("FEDI" & " " & strChan)
            ElseIf strNote = "繳稅" Then
                strChan = Trim("繳稅" & " " & strChan)
            ElseIf strNote = "一卡通" Then
                strChan = Trim("一卡通" & " " & strChan)
            ElseIf Len(strNote) < 4 And Utils.ContainsNumbers(strNote) = True Then
                strChan = "跨行機台"        ' (" & Format(CInt(strNote), "000") & ")"
            End If

            ' FIXME: 轉出入帳號是不是都Ｖ打頭就是簽帳卡交易？　=IF(OR(COUNTIF(J9,"Ｖ*"),COUNTIF(K9,"Ｖ*")),"簽帳卡交易","")
        End If

        If strClerk = "99998" Or strClerk = "99997" Then
            strChan = Trim(strChan & " 系統自動")
        ElseIf strStore <> "0880" And dictBranchName.Exists(strStore) Then
            strChan = Trim(strChan & " 臨櫃分行")
        ElseIf strSummary = "電匯" Then
            strChan = Trim(strChan & " 電匯")
        End If

        wsOut.Cells(i, "AB").Value = strChan

        ' "TranType"
        If strStore = "" Then
            wsOut.Cells(i, "T").Value = ""
        Else
            'FIXME: should have beter way to do it
            If strStore = "0880" Then
                wsOut.Cells(i, "T").Value = "網路或ATM"
            Else
                If Not (strClerk = "99998" Or strClerk = "99997") Then        ' 非自動化通路
                    If dictBranchName.Exists(strStore) Then
                        wsOut.Cells(i, "T").Value = "臨櫃分行"
                        wsOut.Cells(i, "X").Value = dictBranchName(strStore) & "分行"
                    Else
                        wsOut.Cells(i, "T").Value = ""
                        wsOut.Cells(i, "X").Value = "未知分行" & strStore
                    End If

                    If dictBranchCity.Exists(strStore) Then
                        wsOut.Cells(i, "Y").Value = Left(dictBranchCity(strStore), 3)
                        wsOut.Cells(i, "Z").Value = Mid(dictBranchCity(strStore), 4)
                        wsOut.Cells(i, "AA").Value = dictBranchCity(strStore)

                    Else
                        wsOut.Cells(i, "Y").Value = ""        ' wsOut.Cells(i, "F").Value FIXME:how to handle not found
                        wsOut.Cells(i, "Z").Value = ""
                    End If
                End If

            End If
        End If

        ' ATM related
        If strClerk <> "" Then
            If dictATMLoc.Exists(strClerk) Then
                wsOut.Cells(i, "U").Value = dictATMLoc(strClerk)
            Else
                wsOut.Cells(i, "U").Value = ""        ' wsOut.Cells(i, "F").Value
            End If

            If dictATMCity.Exists(strClerk) Then
                wsOut.Cells(i, "V").Value = Left(dictATMCity(strClerk), 3)
                wsOut.Cells(i, "W").Value = Mid(dictATMCity(strClerk), 4)
                wsOut.Cells(i, "AA").Value = dictATMCity(strClerk)

            Else
                wsOut.Cells(i, "V").Value = ""        ' wsOut.Cells(i, "F").Value FIXME:how to handle not found
                wsOut.Cells(i, "W").Value = ""        '
            End If

        End If

GoSkip:
    Next i

    For i = 1 To 27
        wsOut.columns(i).AutoFit
    Next i
    wsOut.Rows(8).HorizontalAlignment = xlCenter
    wsOut.Range("A8:AA" & CStr(lastRow)).AutoFilter

    'freeze
    wsOut.Activate
    With ActiveWindow
        If .FreezePanes Then .FreezePanes = False
        .SplitRow = 8
        .FreezePanes = True
    End With

    Range(posRunStatus).Value = "整理完畢"

End Sub

Sub CreatePivotTables()
    Dim ws          As Worksheet
    Dim wsData      As Worksheet
    Dim dataRange   As Range
    Dim pivotTableRange As Range
    Dim PivotTable  As PivotTable
    Dim pivotCache  As pivotCache
    Dim rowCount    As Long
    Dim rowCurrent  As Long
    Dim rowGap      As Long
    Dim rowPivot    As Long
    Dim offsetiPvot As Long

    Dim i           As Long
    Dim pos         As Integer

    Dim pitems      As Object
    Dim colCityNames As Variant
    colCityNames = Array("基隆市", "新北市", "台北市", "桃園市", "新竹縣", "新竹市", "苗栗縣", "台中市", _
                   "彰化縣", "南投縣", "雲林縣", "嘉義縣", "嘉義市", "台南市", "高雄市", "屏東縣", _
                   "宜蘭縣", "花蓮縣", "台東縣", "澎湖縣", "連江縣")

    rowCurrent = 3
    rowGap = 6
    offsetiPvot = 8
    ' 設定工作表

    Set ws = ThisWorkbook.Sheets(SheetNameMoney)
    Set wsData = ThisWorkbook.Sheets(SheetNameInputData)

    ' 刪除現有的樞紐分析表（如果存在）
    For Each PivotTable In ws.PivotTables
        PivotTable.TableRange2.Clear
        '   pivotTable.pivotCache.Clear
    Next PivotTable

    ' 設定資料範圍
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).row

    Set dataRange = wsData.Range("A8:AC" & CStr(lastRow))

    ' Cash flow related PivotTables
    ' MyPivotTable1 ===================================
    ' Setup 存入交易樞紐分析表
    ' 設定樞紐分析表的位置
    Set pivotTableRange = ws.Range("A" & CStr(rowCurrent))

    ' 設定樞紐分析表的快取範圍
    Set pivotCache = ActiveWorkbook.PivotCaches.Create( _
    SourceType:=xlDatabase, _
    SourceData:=dataRange)

    ' 在指定位置插入樞紐分析表
    Set PivotTable = pivotCache.CreatePivotTable( _
    TableDestination:=pivotTableRange, _
    TableName:="MyPivotTable1")

    ' 在樞紐分析表中設定行、列、值
    With PivotTable
        .CompactLayoutColumnHeader = "交易模式"
        .CompactLayoutRowHeader = "年月"

        .PivotFields("交易月份").Orientation = xlRowField
        .PivotFields("交易摘要").Orientation = xlColumnField
        .PivotFields("In").Orientation = xlDataField

        With .PivotFields("In")
            .Orientation = xlDataField
            .Function = xlSum        ' 設定聚合函數
            .NumberFormat = "#,##0.00"        ' 設定數值格式
        End With

        With ws.Range("A" & CStr(.TableRange1.Cells(1, 1).row - 1))
            .Value = "存入交易"
            .Font.color = ColorWhite
            .Font.Bold = True
            .Interior.color = ColorGreen
        End With

        ' Move to next position
        rowCount = .TableRange1.Rows.Count
        rowCurrent = rowCurrent + rowCount + rowGap
    End With
    ' =================================================

    ' MyPivotTable2 ===================================
    ' Setup 支出交易樞紐分析表
    ' 設定樞紐分析表的位置
    Set pivotTableRange = ws.Range("A" & CStr(rowCurrent))

    ' 設定樞紐分析表的快取範圍
    Set pivotCache = ActiveWorkbook.PivotCaches.Create( _
    SourceType:=xlDatabase, _
    SourceData:=dataRange)

    ' 在指定位置插入樞紐分析表
    Set PivotTable = pivotCache.CreatePivotTable( _
    TableDestination:=pivotTableRange, _
    TableName:="MyPivotTable2")

    ' 在樞紐分析表中設定行、列、值
    With PivotTable
        .CompactLayoutColumnHeader = "交易模式"
        .CompactLayoutRowHeader = "年月"

        .PivotFields("交易月份").Orientation = xlRowField
        .PivotFields("交易摘要").Orientation = xlColumnField
        .PivotFields("Out").Orientation = xlDataField

        With .PivotFields("Out")
            .Orientation = xlDataField
            .Function = xlSum        ' 設定聚合函數
            .NumberFormat = "#,##0.00"        ' 設定數值格式
        End With

        With ws.Range("A" & CStr(.TableRange1.Cells(1, 1).row - 1))
            .Value = "支出交易"
            .Font.color = ColorWhite
            .Font.Bold = True
            .Interior.color = ColorGreen
        End With

        ' Move to next position
        rowCount = .TableRange1.Rows.Count
        rowCurrent = rowCurrent + rowCount + rowGap
    End With
    ' =================================================

    ' MyPivotTable3 ===================================
    Set pivotTableRange = ws.Range("A" & CStr(rowCurrent))

    ' 設定樞紐分析表的快取範圍
    Set pivotCache = ActiveWorkbook.PivotCaches.Create( _
    SourceType:=xlDatabase, _
    SourceData:=dataRange)

    ' 在指定位置插入樞紐分析表
    Set PivotTable = pivotCache.CreatePivotTable( _
    TableDestination:=pivotTableRange, _
    TableName:="MyPivotTable3")

    ' 在樞紐分析表中設定行、列、值
    With PivotTable
        ' .CompactLayoutColumnHeader = "時間"
        ' .CompactLayoutRowHeader = "年月"

        .PivotFields("交易摘要").Orientation = xlPageField
        .PivotFields("轉出入帳號").Orientation = xlRowField

        .AddDataField .PivotFields("In"), "計數 - In", xlCount
        .AddDataField .PivotFields("In"), "加總 - In2", xlSum
        .AddDataField .PivotFields("In"), "佔比 - In3", xlSum

        .PivotFields("佔比 - In3").Calculation = xlPercentOfTotal
        .PivotFields("佔比 - In3").NumberFormat = "0%"
        .PivotFields("交易摘要").ClearAllFilters

        ' 有這個 item 才用這個 item filter
        bExisted04 = ExcelUtils.CheckPivotItemExisted(PivotTable, "交易摘要", "04本行轉帳")
        If bExisted04 = True Then
            .PivotFields("交易摘要").CurrentPage = "04本行轉帳"
        End If

        .PivotFields("轉出入帳號").PivotFilters.Add2 Type _
                                                :=xlValueIsGreaterThan, DataField:=.PivotFields("加總 - In2"), Value1:=0
        .PivotFields("轉出入帳號").AutoSort xlDescending, "加總 - In2"

        ' Keep top k only
        With .TableRange1
            .Rows.Hidden = False
            If .Rows.Count > offsetiPvot Then
                ' 把超過 offsetiPvot 的 row 隱藏起來
                .Rows(CStr(offsetiPvot) & ":" & CStr(.Rows.Count - 1)).Hidden = True
            End If
        End With

        rowPivot = .TableRange1.Cells(1, 1).row

        ws.Range("E" & CStr(rowPivot + 1)).Value = "姓名"
        ws.Range("F" & CStr(rowPivot + 1)).Value = "id"
        ws.Range("G" & CStr(rowPivot + 1)).Value = "職業背景"
        ws.Range("H" & CStr(rowPivot + 1)).Value = "高風險"
        ws.Range("I" & CStr(rowPivot + 1)).Value = "與客戶關係"
        ws.Range("J" & CStr(rowPivot + 1)).Value = "備註"

        With ws.Range("E" & CStr(rowPivot + 1) & ":J" & CStr(rowPivot + 1))
            .Font.color = ColorWhite
            .Font.Bold = True
            .Interior.color = ColorGreen
        End With

        ws.Range("E" & CStr(rowPivot + 2) & ":J" & CStr(rowPivot + 6)).Interior.color = ColorYellow

        With ws.Range("E" & CStr(rowPivot + 1) & ":J" & CStr(rowPivot + 6)).Borders
            .LineStyle = xlContinuous
            .color = RGB(153, 188, 133)        ' Green border color
            .TintAndShade = 0
            .Weight = xlThin
        End With

        With ws.Range("A" & CStr(rowPivot - 1))
            .Value = "轉帳存入交易By總流量"
            .Font.color = ColorWhite
            .Font.Bold = True
            .Interior.color = ColorGreen
        End With

        ' Move to next position
        rowCount = .TableRange1.Rows.Count
        rowCurrent = rowCurrent + rowCount + rowGap + 3
    End With
    ' =================================================

    ' MyPivotTable4 ===================================
    ' 設定樞紐分析表的位置
    Set pivotTableRange = ws.Range("A" & CStr(rowCurrent))

    ' 設定樞紐分析表的快取範圍
    Set pivotCache = ActiveWorkbook.PivotCaches.Create( _
    SourceType:=xlDatabase, _
    SourceData:=dataRange)

    ' 在指定位置插入樞紐分析表
    Set PivotTable = pivotCache.CreatePivotTable( _
    TableDestination:=pivotTableRange, _
    TableName:="MyPivotTable4")

    ' 在樞紐分析表中設定行、列、值
    With PivotTable
        ' .CompactLayoutColumnHeader = "時間"
        ' .CompactLayoutRowHeader = "年月"

        .PivotFields("交易摘要").Orientation = xlPageField
        .PivotFields("轉出入帳號").Orientation = xlRowField

        .AddDataField .PivotFields("In"), "計數 - In", xlCount
        .AddDataField .PivotFields("In"), "加總 - In2", xlSum
        .AddDataField .PivotFields("In"), "佔比 - In3", xlSum

        .PivotFields("佔比 - In3").Calculation = xlPercentOfTotal
        .PivotFields("佔比 - In3").NumberFormat = "0%"
        .PivotFields("交易摘要").ClearAllFilters

        ' 有這個 item 才用這個 item filter
        bExisted03 = ExcelUtils.CheckPivotItemExisted(PivotTable, "交易摘要", "03跨行轉帳")
        If bExisted03 = True Then
            .PivotFields("交易摘要").CurrentPage = "03跨行轉帳"
        End If

        .PivotFields("轉出入帳號").PivotFilters.Add2 Type _
                                                :=xlValueDoesNotEqual, DataField:=.PivotFields("加總 - In2"), Value1:=0

        .PivotFields("轉出入帳號").AutoSort xlDescending, "加總 - In2"

        rowPivot = .TableRange1.Cells(1, 1).row

        ws.Range("E" & CStr(rowPivot + 1)).Value = "姓名"
        ws.Range("F" & CStr(rowPivot + 1)).Value = "id"

        With ws.Range("E" & CStr(rowPivot + 1) & ":F" & CStr(rowPivot + 1))
            .Font.color = ColorWhite
            .Font.Bold = True
            .Interior.color = ColorGreen
        End With

        ws.Range("E" & CStr(rowPivot + 2) & ":F" & CStr(rowPivot + 6)).Interior.color = ColorYellow

        With ws.Range("E" & CStr(rowPivot + 1) & ":F" & CStr(rowPivot + 6)).Borders
            .LineStyle = xlContinuous
            .color = RGB(153, 188, 133)        ' Green border color
            .TintAndShade = 0
            .Weight = xlThin
        End With

        With ws.Range("A" & CStr(rowPivot - 1))
            .Value = "跨行轉帳存入交易By總流量"
            .Font.color = ColorWhite
            .Font.Bold = True
            .Interior.color = ColorGreen
        End With

        ' Move to next position
        rowCount = .TableRange1.Rows.Count
        rowCurrent = rowCurrent + rowCount + rowGap + 3

        ' Keep top k only
        With .TableRange1
            .Rows.Hidden = False
            If .Rows.Count > offsetiPvot Then
                ' 把超過 offsetiPvot 的 row 隱藏起來
                .Rows(CStr(offsetiPvot) & ":" & CStr(.Rows.Count - 1)).Hidden = True
            End If
        End With
    End With
    ' =================================================

    ' MyPivotTable5 ===================================
    Set pivotTableRange = ws.Range("A" & CStr(rowCurrent))
    ' 設定樞紐分析表的快取範圍

    Set pivotCache = ActiveWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)

    ' 在指定位置插入樞紐分析表
    Set PivotTable = pivotCache.CreatePivotTable( _
    TableDestination:=pivotTableRange, _
    TableName:="MyPivotTable5")

    ' 在樞紐分析表中設定行、列、值
    With PivotTable
        ' .CompactLayoutColumnHeader = "時間"
        ' .CompactLayoutRowHeader = "年月"

        .PivotFields("交易摘要").Orientation = xlPageField
        .PivotFields("轉出入帳號").Orientation = xlRowField

        .AddDataField .PivotFields("Out"), "計數 - Out", xlCount
        .AddDataField .PivotFields("Out"), "加總 - Out2", xlSum
        .AddDataField .PivotFields("Out"), "佔比 - Out3", xlSum

        .PivotFields("佔比 - Out3").Calculation = xlPercentOfTotal
        .PivotFields("佔比 - Out3").NumberFormat = "0%"
        .PivotFields("交易摘要").ClearAllFilters
        .PivotFields("交易摘要").CurrentPage = "04本行轉帳"

        .PivotFields("轉出入帳號").PivotFilters.Add2 Type _
                                                :=xlValueIsGreaterThan, DataField:=.PivotFields("加總 - Out2"), Value1:=0

        .PivotFields("轉出入帳號").AutoSort xlDescending, "加總 - Out2"

        rowPivot = .TableRange1.Cells(1, 1).row

        ws.Range("E" & CStr(rowPivot + 1)).Value = "姓名"
        ws.Range("F" & CStr(rowPivot + 1)).Value = "id"
        ws.Range("G" & CStr(rowPivot + 1)).Value = "職業背景"
        ws.Range("H" & CStr(rowPivot + 1)).Value = "高風險"
        ws.Range("I" & CStr(rowPivot + 1)).Value = "與客戶關係"
        ws.Range("J" & CStr(rowPivot + 1)).Value = "備註"

        With ws.Range("E" & CStr(rowPivot + 1) & ":J" & CStr(rowPivot + 1))
            .Font.color = ColorWhite
            .Font.Bold = True
            .Interior.color = ColorGreen
        End With

        ws.Range("E" & CStr(rowPivot + 2) & ":J" & CStr(rowPivot + 6)).Interior.color = ColorYellow

        With ws.Range("E" & CStr(rowPivot + 1) & ":J" & CStr(rowPivot + 6)).Borders
            .LineStyle = xlContinuous
            .color = RGB(153, 188, 133)        ' Green border color
            .TintAndShade = 0
            .Weight = xlThin
        End With

        With ws.Range("A" & CStr(rowPivot - 1))
            .Value = "轉帳支出交易By總流量"
            .Font.color = ColorWhite
            .Font.Bold = True
            .Interior.color = ColorGreen
        End With

        ' Keep top k only
        With .TableRange1
            .Rows.Hidden = False
            If .Rows.Count > offsetiPvot Then
                ' 把超過 offsetiPvot 的 row 隱藏起來
                .Rows(CStr(offsetiPvot) & ":" & CStr(.Rows.Count - 1)).Hidden = True
            End If
        End With

        ' Move to next position
        rowCount = .TableRange1.Rows.Count
        rowCurrent = rowCurrent + rowCount + rowGap + 3
    End With
    ' =================================================

    ' MyPivotTable6 ===================================
    Set pivotTableRange = ws.Range("A" & CStr(rowCurrent))

    ' 設定樞紐分析表的快取範圍
    Set pivotCache = ActiveWorkbook.PivotCaches.Create( _
    SourceType:=xlDatabase, _
    SourceData:=dataRange)

    ' 在指定位置插入樞紐分析表
    Set PivotTable = pivotCache.CreatePivotTable( _
    TableDestination:=pivotTableRange, _
    TableName:="MyPivotTable6")

    ' 在樞紐分析表中設定行、列、值
    With PivotTable
        ' .CompactLayoutColumnHeader = "時間"
        ' .CompactLayoutRowHeader = "年月"

        .PivotFields("交易摘要").Orientation = xlPageField
        .PivotFields("轉出入帳號").Orientation = xlRowField

        .AddDataField .PivotFields("Out"), "計數 - Out", xlCount
        .AddDataField .PivotFields("Out"), "加總 - Out2", xlSum
        .AddDataField .PivotFields("Out"), "佔比 - Out3", xlSum

        .PivotFields("佔比 - Out3").Calculation = xlPercentOfTotal
        .PivotFields("佔比 - Out3").NumberFormat = "0%"
        .PivotFields("交易摘要").ClearAllFilters

        ' 有這個 item 才用這個 item filter
        bExisted03 = ExcelUtils.CheckPivotItemExisted(PivotTable, "交易摘要", "03跨行轉帳")
        If bExisted03 = True Then
            .PivotFields("交易摘要").CurrentPage = "03跨行轉帳"
        End If

        .PivotFields("轉出入帳號").PivotFilters.Add2 Type _
                                                :=xlValueDoesNotEqual, DataField:=.PivotFields("加總 - Out2"), Value1:=0

        .PivotFields("轉出入帳號").AutoSort xlDescending, "加總 - Out2"

        rowPivot = .TableRange1.Cells(1, 1).row

        ws.Range("E" & CStr(rowPivot + 1)).Value = "姓名"
        ws.Range("F" & CStr(rowPivot + 1)).Value = "id"

        With ws.Range("E" & CStr(rowPivot + 1) & ":F" & CStr(rowPivot + 1))
            .Font.color = ColorWhite
            .Font.Bold = True
            .Interior.color = ColorGreen
        End With

        ws.Range("E" & CStr(rowPivot + 2) & ":F" & CStr(rowPivot + 6)).Interior.color = ColorYellow

        With ws.Range("E" & CStr(rowPivot + 1) & ":F" & CStr(rowPivot + 6)).Borders
            .LineStyle = xlContinuous
            .color = RGB(153, 188, 133)        ' Green border color
            .TintAndShade = 0
            .Weight = xlThin
        End With

        With ws.Range("A" & CStr(rowPivot - 1))
            .Value = "跨行轉帳支出交易By總流量"
            .Font.color = ColorWhite
            .Font.Bold = True
            .Interior.color = ColorGreen
        End With

        ' Keep top k only
        With .TableRange1
            .Rows.Hidden = False
            If .Rows.Count > offsetiPvot Then
                ' 把超過 offsetiPvot 的 row 隱藏起來
                .Rows(CStr(offsetiPvot) & ":" & CStr(.Rows.Count - 1)).Hidden = True
            End If
        End With

        ' Move to next position
        rowCount = .TableRange1.Rows.Count
        rowCurrent = rowCurrent + rowCount + rowGap + 6
    End With
    ' =================================================

    ' ATM related PivotTables
    ' 用 gCntMultiATM 以及 gCntMultiDeposit 兩個值來判斷是否要畫出 ATM 相關的 PivotTable
    If gCntMultiATM > 0 Or gCntMultiDeposit > 0 Then

        ' MyPivotTable7 ===================================
        ' Setup ATM機台存入樞紐分析表

        ' 設定樞紐分析表的位置
        Set pivotTableRange = ws.Range("A" & CStr(rowCurrent))

        ' 設定樞紐分析表的快取範圍
        Set pivotCache = ActiveWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)

        ' 在指定位置插入樞紐分析表
        Set PivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=pivotTableRange, _
        TableName:="MyPivotTable7")

        ' 在樞紐分析表中設定行、列、值
        With PivotTable
            .CompactLayoutColumnHeader = "時間"
            .CompactLayoutRowHeader = "年月"

            .PivotFields("ATM縣市").Orientation = xlPageField
            .PivotFields("交易摘要").Orientation = xlPageField
            .PivotFields("交易月份").Orientation = xlRowField
            .PivotFields("In").Orientation = xlDataField
            .PivotFields("簡化交易時間").Orientation = xlColumnField
            .PivotFields("交易摘要").ClearAllFilters

            ' 有這個 item 才用這個 item filter
            bExisted01 = ExcelUtils.CheckPivotItemExisted(PivotTable, "交易摘要", "01現金交易")
            If bExisted01 = True Then
                .PivotFields("交易摘要").CurrentPage = "01現金交易"
            End If

            With ws.Range("A" & CStr(.TableRange1.Cells(1, 1).row - 1))
                .Value = "ATM存入交易時間"
                .Font.color = ColorWhite
                .Font.Bold = True
                .Interior.color = ColorGreen
            End With

            ' Move to next position
            rowCount = .TableRange1.Rows.Count
            rowCurrent = rowCurrent + rowCount + rowGap
        End With
        ' =================================================

        ' MyPivotTable8 ===================================
        ' Setup ATM機台存入樞紐分析表

        ' 設定樞紐分析表的位置
        Set pivotTableRange = ws.Range("A" & CStr(rowCurrent))

        ' 設定樞紐分析表的快取範圍
        Set pivotCache = ActiveWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)

        ' 在指定位置插入樞紐分析表
        Set PivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=pivotTableRange, _
        TableName:="MyPivotTable8")

        ' 在樞紐分析表中設定行、列、值
        With PivotTable
            .CompactLayoutColumnHeader = "縣市"
            .CompactLayoutRowHeader = "年月"

            .PivotFields("ATM縣市").Orientation = xlColumnField
            .PivotFields("交易摘要").Orientation = xlPageField
            .PivotFields("交易月份").Orientation = xlRowField
            .PivotFields("In").Orientation = xlDataField
            .PivotFields("交易摘要").ClearAllFilters

            ' 有這個 item 才用這個 item filter
            bExisted01 = ExcelUtils.CheckPivotItemExisted(PivotTable, "交易摘要", "01現金交易")
            If bExisted01 = True Then
                .PivotFields("交易摘要").CurrentPage = "01現金交易"
            End If

            Set pitems = .PivotFields("ATM縣市").PivotItems
            pos = 2
            If pitems.Count > 0 And pitems(1).name = "" Then
                pitems(1).name = "他行ATM"
            End If

            For i = LBound(colCityNames) To UBound(colCityNames)
                If Utils.ObjectContainsItem(pitems, colCityNames(i)) = True Then
                    pitems(colCityNames(i)).position = pos
                    pos = pos + 1
                End If
            Next i

            With ws.Range("A" & CStr(.TableRange1.Cells(1, 1).row - 1))
                .Value = "ATM存入交易地點"
                .Font.color = ColorWhite
                .Font.Bold = True
                .Interior.color = ColorGreen
            End With

            ' Move to next position
            rowCount = .TableRange1.Rows.Count
            rowCurrent = rowCurrent + rowCount + rowGap + 3
        End With
        ' =================================================

        ' MyPivotTable9 ===================================
        ' Setup ATM機台支出樞紐分析表

        ' 設定樞紐分析表的位置
        Set pivotTableRange = ws.Range("A" & CStr(rowCurrent))

        ' 設定樞紐分析表的快取範圍
        Set pivotCache = ActiveWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)

        ' 在指定位置插入樞紐分析表
        Set PivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=pivotTableRange, _
        TableName:="MyPivotTable9")

        ' 在樞紐分析表中設定行、列、值
        With PivotTable
            .CompactLayoutColumnHeader = "時間"
            .CompactLayoutRowHeader = "年月"

            .PivotFields("ATM縣市").Orientation = xlPageField
            .PivotFields("交易摘要").Orientation = xlPageField
            .PivotFields("交易月份").Orientation = xlRowField
            .PivotFields("Out").Orientation = xlDataField
            .PivotFields("簡化交易時間").Orientation = xlColumnField
            .PivotFields("交易摘要").ClearAllFilters

            ' 有這個 item 才用這個 item filter
            bExisted01 = ExcelUtils.CheckPivotItemExisted(PivotTable, "交易摘要", "01現金交易")
            If bExisted01 = True Then
                .PivotFields("交易摘要").CurrentPage = "01現金交易"
            End If

            With ws.Range("A" & CStr(.TableRange1.Cells(1, 1).row - 1))
                .Value = "ATM領出交易時間"
                .Font.color = ColorWhite
                .Font.Bold = True
                .Interior.color = ColorGreen
            End With

            ' Move to next position
            rowCount = .TableRange1.Rows.Count
            rowCurrent = rowCurrent + rowCount + rowGap
        End With
        ' =================================================

        ' MyPivotTable10 ===================================
        ' Setup ATM機台存入樞紐分析表

        ' 設定樞紐分析表的位置
        Set pivotTableRange = ws.Range("A" & CStr(rowCurrent))

        ' 設定樞紐分析表的快取範圍
        Set pivotCache = ActiveWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)

        ' 在指定位置插入樞紐分析表
        Set PivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=pivotTableRange, _
        TableName:="MyPivotTable10")

        ' 在樞紐分析表中設定行、列、值
        With PivotTable
            .CompactLayoutColumnHeader = "縣市"
            .CompactLayoutRowHeader = "年月"

            .PivotFields("ATM縣市").Orientation = xlColumnField
            .PivotFields("交易摘要").Orientation = xlPageField
            .PivotFields("交易月份").Orientation = xlRowField
            .PivotFields("Out").Orientation = xlDataField
            .PivotFields("交易摘要").ClearAllFilters

            ' 有這個 item 才用這個 item filter
            bExisted01 = ExcelUtils.CheckPivotItemExisted(PivotTable, "交易摘要", "01現金交易")
            If bExisted01 = True Then
                .PivotFields("交易摘要").CurrentPage = "01現金交易"
            End If

            Set pitems = .PivotFields("ATM縣市").PivotItems
            pos = 2
            If pitems.Count > 0 And pitems(1).name = "" Then
                pitems(1).name = "其他"
            End If

            For i = LBound(colCityNames) To UBound(colCityNames)
                If Utils.ObjectContainsItem(pitems, colCityNames(i)) = True Then
                    pitems(colCityNames(i)).position = pos
                    pos = pos + 1
                End If
            Next i

            With ws.Range("A" & CStr(.TableRange1.Cells(1, 1).row - 1))
                .Value = "ATM領出交易地點"
                .Font.color = ColorWhite
                .Font.Bold = True
                .Interior.color = ColorGreen
            End With

            ' Move to next position
            rowCount = .TableRange1.Rows.Count
            rowCurrent = rowCurrent + rowCount + rowGap
        End With
        ' =================================================

    End If
    ' End of ATM related PivotTables

    ExcelUtils.SetSheetDefStyle SheetNameMoney
    ws.Activate
    ActiveWindow.ScrollRow = 1
End Sub

Function GetKVPairsToDictOld(wsRef As Worksheet, colKey As String, colVal As String) As Object
    'Dim wsRef      As Worksheet
    Dim dict        As Object
    Dim lastRow     As Long
    Dim i           As Long
    Dim item        As Variant

    Set dict = CreateObject("Scripting.Dictionary")

    lastRow = wsRef.Cells(wsRef.Rows.Count, colKey).End(xlUp).row
    ' Loop through column G and add items to the dictionary
    For i = 1 To lastRow
        item = CStr(wsRef.Cells(i, colKey).Value)
        If Not IsEmpty(item) Then
            If Not dict.Exists(item) Then
                'Debug.Print "add blacklist item: " & blacklistItem
                dict.Add item, wsRef.Cells(i, colVal).Value
            End If
        End If
    Next i

    ' You can use the dictionary as needed
    Set GetKVPairsToDict = dict

End Function

Function IsThresholdArrayEmpty(arr As Variant) As Boolean
    If UBound(arr) = 1 And arr(1) = EmptyArrayValue Then
        IsThresholdArrayEmpty = True
    Else
        IsThresholdArrayEmpty = False
    End If
End Function

' Create bar charts in page Main
Sub CreateBarChart(rowIndex1 As Integer, rowIndex2 As Integer, targetPosition As Range, caption As String, yValue As String)
    Dim ws          As Worksheet
    Dim chartObj    As chartObject
    Dim lastCol1    As Integer, lastCol2 As Integer
    Dim dataRange1  As Range, dataRange2 As Range
    Dim xAxisRange  As Range, yAxisRange As Range
    Dim chartRange  As Range
    Dim wsOut       As Worksheet
    ' Set the worksheet
    Set ws = Worksheets(SheetNameIntermediate)
    Set wsOut = Worksheets(SheetNameMain)
    ' Find the last column in each row
    lastCol1 = ws.Cells(rowIndex1, ws.columns.Count).End(xlToLeft).Column
    ' Assume 跟col1 一樣
    lastCol2 = lastCol1
    ' Define the data ranges For the rows
    Set dataRange1 = ws.Range(ws.Cells(rowIndex1, 1), ws.Cells(rowIndex1, lastCol1))
    Set dataRange2 = ws.Range(ws.Cells(rowIndex2, 1), ws.Cells(rowIndex2, lastCol2))
    ' Combine the data ranges into one range
    Set chartRange = Union(dataRange1, dataRange2)
    Set xAxisRange = dataRange1
    Set yAxisRange = dataRange2
    ' Create a New chart on the worksheet
    Set chartObj = wsOut.ChartObjects.Add(Left:=targetPosition.Left, Top:=targetPosition.Top, Width:=targetPosition.Width, Height:=targetPosition.Height)
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
End Sub
