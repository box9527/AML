' =================================================
' Sheet page 3.1交易明細
' =================================================

Option Explicit

' Coding rules
' 1. enable "Option Explicit"
' 2. use type Long to reference rows and cols
' 3. use type "Double" to store amount of money
' ref: https://learn.microsoft.com/zh-tw/office/vba/language/reference/user-interface-help/data-type-summary
' 4. do not use type "Integer"
' 5. layout specific constants (e.g. names of rows / columns)
' 6. No magic numbers
' 7. Not to use type "Variant" if possible
' 8. Dbg.Print slows down the program; using Excel native functions slows down the program (the switching takes time)
' 9. Check the data type before using it. (the data may not be as expected)
' 10. Put enough information in debugging logs

' Define layout specific constants here
' 定義畫在最頂層的進階選項
Private Const chkEnableBalance   As String = "P1"
Private Const chkEnableTxnTime   As String = "P2"
Private Const chkEnableLarge     As String = "P3"
Private Const chkEnableNearLarge As String = "P4"
Private Const chkUnique          As String = "P5"
Private Const txtSomeAcc         As String = "Q1"

Private Const cbWin          As String = "R1"
Private Const cbOccur        As String = "R2"
Private Const cbOppo         As String = "R3"
Private Const chkBranch      As String = "S1"
Private Const chkATM         As String = "S2"
Private Const chkAuto        As String = "S3"
Private Const chkDeposit     As String = "T1"
Private Const chkWithdraw    As String = "T2"
Private Const chkCash        As String = "T3"
Private Const txtAmtMin      As String = "U1"
Private Const txtAmtMax      As String = "U2"
Private Const txtAmtDf       As String = "U3"
Private Const txtAmtLessThan As String = "U4"
Private Const txtAmtTotal    As String = "U5"
Private Const cbLabel        As String = "V1"

Private Const strCannotFindCfg  As String = "無法在自訂標示設定頁找到對應設定:"
Private Const strNeedPickInOut  As String = "當差值數字不等於0時, [存入]與[提出]必須被勾選"
Private Const strWarnTimeWindow As String = "錯誤的時間區間設定"
Private Const strWarnOccurrence As String = "錯誤的頻次設定"
Private Const strDefaultAccount As String = "self"

' Row index in config sheet
Private Const RowInterval    As String = "B"
Private Const RowTimes       As String = "C"
Private Const RowOppo        As String = "D"
Private Const RowUnique      As String = "E"
Private Const RowSomeAcc     As String = "F"
Private Const RowBranch      As String = "G"
Private Const RowATM         As String = "H"
Private Const RowAuto        As String = "I"
Private Const RowDeposit     As String = "J"
Private Const RowWithdraw    As String = "K"
Private Const RowCash        As String = "L"
Private Const RowAmtMin      As String = "M"
Private Const RowAmtMax      As String = "N"
Private Const RowAmtDf       As String = "O"
Private Const RowAmtLessThan As String = "P"
Private Const RowAmtTotal    As String = "Q"
Private Const RowDesc        As String = "R"

Dim gCntFrames As Long
Dim gLabelCur  As String

Private Sub cbPattern_Change()
    Debug.Print "Hit change" & Range(cbLabel)
    RefreshSearchSetting Range(cbLabel)
End Sub

Private Function IsTargetChannels(useBr As Boolean, useATM As Boolean, _
                                  useAuto As Boolean, strChan As String) As Boolean
    Dim isTargetCh As Boolean
    If (((useBr = True) And ((strChan = UiCondiChBranch) Or (strChan = UiCondiChwireTS))) Or _
        ((useATM = True) And ((strChan = UiCondiChATMAuto) Or (strChan = UiCondiChBrDevices))) Or _
        ((useAuto = True) And ((strChan = UiCondiChMobile) Or (strChan = UiCondiChOnline)))) Then
        isTargetCh = True
    End If

    IsTargetChannels = isTargetCh
End Function

Private Sub RefreshSearchSetting(strLabel)
    Dim wsCfg As Worksheet

    ' Find the row
    Dim rowCfg  As Long
    Dim lastRow As Long
    Dim i       As Long

    Set wsCfg = Worksheets(SheetNameLabel)
    lastRow = wsCfg.Cells(wsCfg.Rows.Count, 1).End(xlUp).row

    If strLabel = UiDisableSearch Then
        gLabelCur = Range(cbLabel).value
        Exit Sub
    End If

    rowCfg = 1

    For i = 2 To lastRow
        If wsCfg.Cells(i, 1).value = strLabel Then
            rowCfg = i
            Exit For
        End If
    Next i

    If rowCfg = 1 Then
        MsgBox strCannotFindCfg & strLabel
        Range(cbLabel).value = gLabelCur
        'FIXME: should update pattern description here
        Exit Sub
    End If

    ' Fill row value to UI position
    Range(cbWin).value = wsCfg.Cells(rowCfg, RowInterval)
    Range(cbOccur).value = wsCfg.Cells(rowCfg, RowTimes)
    Range(cbOppo).value = wsCfg.Cells(rowCfg, RowOppo)
    Range(chkUnique).value = wsCfg.Cells(rowCfg, RowUnique)
    Range(txtSomeAcc).value = wsCfg.Cells(rowCfg, RowSomeAcc)
    Range(chkBranch).value = wsCfg.Cells(rowCfg, RowBranch)
    Range(chkATM).value = wsCfg.Cells(rowCfg, RowATM)
    Range(chkAuto).value = wsCfg.Cells(rowCfg, RowAuto)
    Range(chkDeposit).value = wsCfg.Cells(rowCfg, RowDeposit)
    Range(chkWithdraw).value = wsCfg.Cells(rowCfg, RowWithdraw)
    Range(chkCash).value = wsCfg.Cells(rowCfg, RowCash)
    Range(txtAmtMin).value = wsCfg.Cells(rowCfg, RowAmtMin)
    Range(txtAmtMax).value = wsCfg.Cells(rowCfg, RowAmtMax)
    Range(txtAmtDf).value = wsCfg.Cells(rowCfg, RowAmtDf)
    Range(txtAmtLessThan).value = wsCfg.Cells(rowCfg, RowAmtLessThan)
    Range(txtAmtTotal).value = wsCfg.Cells(rowCfg, RowAmtTotal)

    Worksheets(SheetNameSimple).OLEObjects("ptnDesc").Object.caption = wsCfg.Cells(rowCfg, RowDesc).value

    gLabelCur = Range(cbLabel).value
End Sub

Private Sub Refresh_Click()
    Refresh_Click_Int True
    ExcelUtils.ApplySheetSimpleStyle
End Sub

Private Sub Refresh_Click_Int(blShowOk As Boolean)
    Dim ws           As Worksheet
    Dim enableSearch As Boolean

    Set ws = ThisWorkbook.Sheets(SheetNameSimple)

    If Range(cbLabel).value <> UiDisableSearch Then
        enableSearch = True
    End If

    gCntFrames = 0
    Application.ScreenUpdating = False        ' for performance

    ' Reset all cells
    ' 把所有的格內容清空
    ResetPage ws
    ' 逐列檢查資料
    CheckOneByOne ws

    If enableSearch = True Then
        ' 依時間區間聚集
        TraverseRowsByTimeWindow ws
    End If

    Application.ScreenUpdating = True

    ' 有要求顯示完成對話框的話
    If blShowOk = True Then
        MsgBox BtnStatusStopRebase
    End If
End Sub

Private Sub CheckOneByOne(ByVal ws As Worksheet)
    Dim i           As Long
    Dim lastRow     As Long
    Dim cellTime    As Double
    Dim strTmp      As String
    Dim strAccount  As String
    Dim enableBalance As Boolean
    Dim enableTxnTime As Boolean
    Dim enableLarge   As Boolean
    Dim enableNearLarge As Boolean
    Dim strReason       As String

    enableBalance = Range(chkEnableBalance).value
    enableTxnTime = Range(chkEnableTxnTime).value
    enableLarge = Range(chkEnableLarge).value
    enableNearLarge = Range(chkEnableNearLarge).value

    lastRow = ws.Cells(ws.Rows.Count, ColShSimpleTSTime).End(xlUp).row

    For i = RowDataBegin To lastRow
        ' Check balance
        If enableBalance = True Then
            If VarType(ws.Cells(i, ColShSimpleBalance).value) = vbDouble Then
                If ws.Cells(i, ColShSimpleBalance).value < ColShSimpleBalance1K Then
                    ws.Cells(i, ColShSimpleBalance).Font.color = ColorRed
                End If
            End If
        End If

        ' Check amount
        If (VarType(ws.Cells(i, ColShSimpleTSIn).value) = vbDouble) And (Not IsEmpty(ws.Cells(i, ColShSimpleTSIn))) Then
            If (ws.Cells(i, ColShSimpleTSIn).value < ColShSimpleTSIn500K) And _
               (ws.Cells(i, ColShSimpleTSIn).value >= ColShSimpleTSIn450K) Then
                ws.Cells(i, ColShSimpleTSIn).Font.color = ColorRed
            End If

            If ws.Cells(i, ColShSimpleTSIn).value >= ColShSimpleTSIn500K Then
                ws.Cells(i, ColShSimpleTSIn).Font.color = ColorRed
            End If
        End If

        If (VarType(ws.Cells(i, ColShSimpleTSOut).value) = vbDouble) And (Not IsEmpty(ws.Cells(i, ColShSimpleTSOut))) Then
            If enableNearLarge = True Then
                If (ws.Cells(i, ColShSimpleTSOut).value < ColShSimpleTSOut500K) And _
                   (ws.Cells(i, ColShSimpleTSOut).value >= ColShSimpleTSOut450K) Then
                    ws.Cells(i, ColShSimpleTSOut).Font.color = ColorRed
                End If
            End If

            If enableLarge = True Then
                If ws.Cells(i, ColShSimpleTSOut).value >= ColShSimpleTSOut500K Then
                    ws.Cells(i, ColShSimpleTSOut).Font.color = ColorRed
                End If
            End If
        End If

        ' Check time
        If enableTxnTime = True Then
            If VarType(ws.Cells(i, ColShSimpleTSTime).value) = vbDouble Then
                cellTime = ws.Cells(i, ColShSimpleTSTime).value
                If ExcelUtils.IsMorningTime(cellTime) = True Then
                    ws.Cells(i, ColShSimpleTSTime).Font.color = ColorRed
                End If
            End If
        End If

        ' Check bad accounts
        strReason = ExcelUtils.ConvPivotAccShowName(ws.Cells(i, ColShSimpleTempVAccName).value, _
                                                    ws.Cells(i, ColShSimpleTempWAccName).value)

        ' 標成紅色
        If Len(strReason) > 0 Then
            ws.Cells(i, ColShSimpleAlertReason).value = strReason
            ws.Cells(i, ColShSimpleAlertReason).Font.color = ColorWhite
            ws.Cells(i, ColShSimpleAlertReason).Interior.color = ColorRed

            ws.Cells(i, ColShSimpleTSInOutAcc).Font.color = ColorWhite
            ws.Cells(i, ColShSimpleTSInOutAcc).Interior.color = ColorRed
        End If
    Next i
End Sub

Private Sub TraverseRowsByTimeWindow(ByVal ws As Worksheet)
    Dim lastRow  As Long
    Dim currTime As Date
    Dim winStart As Date
    Dim winEnd   As Date

    Dim threshold   As Double
    Dim winLength   As Double
    Dim timeRange   As Double
    Dim i           As Long
    Dim j           As Long
    Dim timeUnit    As String
    Dim nextStart   As Date
    Dim nextEnd     As Date
    Dim nextRowStart As Long

    Dim minOccur    As Long
    Dim strWindow   As String
    Dim strOccur    As String
    Dim strOpponent As String

    Dim useBranch    As Boolean
    Dim useATM       As Boolean
    Dim useAuto      As Boolean
    Dim typeDeposit  As Boolean
    Dim typeWithdraw As Boolean
    Dim typeCash     As Boolean

    Dim bUnique As Boolean

    Dim amtSingleMax  As Double
    Dim amtSingleMin  As Double
    Dim amtDiff       As Double
    Dim amtBalanceMax As Double
    Dim amtAccMax     As Double

    Dim chkSomeAcc  As Boolean
    Dim strSomeAcc  As String

    Dim chkSingleMax  As Boolean
    Dim chkSingleMin  As Boolean
    Dim chkDiff       As Boolean
    Dim chkBalanceMax As Boolean
    Dim chkAccMax     As Boolean

    strWindow = Range(cbWin).value
    strOccur = Range(cbOccur).value
    strOpponent = Range(cbOppo).value
    useBranch = Range(chkBranch).value
    useATM = Range(chkATM).value
    useAuto = Range(chkAuto).value
    typeDeposit = Range(chkDeposit).value
    typeWithdraw = Range(chkWithdraw).value
    typeCash = Range(chkCash).value

    bUnique = Range(chkUnique).value

    If Trim(Range(txtSomeAcc)) <> "" And Range(cbOppo) = UiSpecifiedAcc Then
        strSomeAcc = Trim(Range(txtSomeAcc))
        chkSomeAcc = True
    End If

    If Trim(Range(txtAmtMin)) <> "" Then
        chkSingleMin = True
        amtSingleMin = Range(txtAmtMin)
    End If

    If Trim(Range(txtAmtMax)) <> "" Then
        chkSingleMax = True
        amtSingleMax = Range(txtAmtMax)
    End If

    If Trim(Range(txtAmtDf)) <> "" Then
        chkDiff = True
        amtDiff = Range(txtAmtDf)
    End If

    If Trim(Range(txtAmtLessThan)) <> "" Then
        chkBalanceMax = True
        amtBalanceMax = Range(txtAmtLessThan)
    End If

    If Trim(Range(txtAmtTotal)) <> "" Then
        chkAccMax = True
        amtAccMax = Range(txtAmtTotal)
    End If

    If (chkDiff = True) And ((typeDeposit = False) Or (typeWithdraw = False)) Then
        MsgBox strNeedPickInOut
        Exit Sub
    End If

    timeUnit = UiTimeUnitDay
    winLength = 1 ' Set the time window in hours
    If strWindow = UiTimeWindow1D Then
        winLength = 1
    ElseIf strWindow = UiTimeWindow3D Then
        winLength = 3
    ElseIf strWindow = UiTimeWindow5D Then
        winLength = 5
    ElseIf strWindow = UiTimeWindow1M Then
        timeUnit = UiTimeUnitMonth
    Else
        MsgBox strWarnTimeWindow & strWindow
        Exit Sub
    End If

    If strOccur = UiOccur1Time Then
        minOccur = 1
    ElseIf strOccur = UiOccur3Time Then
        minOccur = 3
    ElseIf strOccur = UiOccur5Time Then
        minOccur = 5
    ElseIf strOccur = UiOccur10Time Then
        minOccur = 10
    Else
        MsgBox strWarnOccurrence & strOccur
        Exit Sub
    End If

    ' Find the last row with data
    lastRow = ws.Cells(ws.Rows.Count, ColShSimpleTSTime).End(xlUp).row

    Dim dictDeposit  As Object
    Dim dictWithdraw As Object
    Dim dictMarkDep  As Object
    Dim dictMarkWith As Object

    Set dictDeposit = CreateObject("Scripting.Dictionary")
    Set dictWithdraw = CreateObject("Scripting.Dictionary")
    Set dictMarkDep = CreateObject("Scripting.Dictionary")
    Set dictMarkWith = CreateObject("Scripting.Dictionary")

    ' Loop through the rows
    nextStart = ws.Cells(RowDataBegin, ColShSimpleTSDate).MergeArea.Cells(1, 1).value + ws.Cells(RowDataBegin, ColShSimpleTSTime).value
    nextRowStart = RowDataBegin
    i = RowDataBegin
    While i <= lastRow
        ' Get the current date and time; Fetch the real row
        currTime = ws.Cells(i, ColShSimpleTSDate).MergeArea.Cells(1, 1).value + ws.Cells(i, ColShSimpleTSTime).value
        ' Calculate the window boundaries
        If currTime >= nextEnd Then
            winStart = currTime
        Else
            winStart = nextStart
        End If

        If timeUnit = UiTimeUnitHour Then
            winEnd = DateAdd("h", winLength, winStart)
            nextStart = DateAdd("h", 1, winStart)
            nextEnd = nextStart + DateAdd("h", winLength, nextStart)

        ElseIf timeUnit = UiTimeUnitDay Then
            winEnd = DateSerial(year(winStart), month(winStart), Day(winStart) + winLength) + _
                     TimeSerial(0, 0, 0)
            nextStart = DateSerial(year(winStart), month(winStart), Day(winStart) + 1) + _
                     TimeSerial(0, 0, 0)
            nextEnd = DateSerial(year(nextStart), month(nextStart), Day(nextStart) + winLength) + _
                     TimeSerial(0, 0, 0)

        ElseIf timeUnit = UiTimeUnitMonth Then
            winEnd = DateSerial(year(winStart), month(winStart) + winLength, 1) + _
                     TimeSerial(0, 0, 0)
            nextStart = DateSerial(year(winStart), month(winStart) + 1, 1) + _
                     TimeSerial(0, 0, 0)
            nextEnd = DateSerial(year(nextStart), month(nextStart) + winLength, 1) + _
                     TimeSerial(0, 0, 0)
        Else
            Debug.Print "Should Not happen"
        End If

        ' Process items inside the window below
        ' Reset the accumulated amount for each window
        Dim cntTxn  As Long
        Dim cntHit  As Long

        Dim amtDeposit  As Double
        Dim sumDeposit  As Double
        Dim amtWithdraw As Double
        Dim sumWithdraw As Double
        Dim amtBalance  As Double
        Dim amtTxn      As Double
        Dim sumTxn      As Double
        Dim hitTop      As Long
        Dim hitBottom   As Long

        dictDeposit.RemoveAll
        dictWithdraw.RemoveAll
        dictMarkDep.RemoveAll
        dictMarkWith.RemoveAll

        Dim rowStep As Long
        Dim accNo   As String
        For j = i To lastRow

            If Not (ws.Cells(j, ColShSimpleTSSummary).comment Is Nothing) Then
                ws.Cells(j, ColShSimpleTSSummary).comment.Delete
            End If

            If Not (ws.Cells(j, ColShSimpleTSSummary).comment Is Nothing) Then
                ws.Cells(j, ColShSimpleTSSummary).comment.Delete
            End If

            ' Get the date and time of the current row
            Dim rowDateTime As Date
            rowDateTime = ws.Cells(j, ColShSimpleTSDate).MergeArea.Cells(1, 1).value + ws.Cells(j, ColShSimpleTSTime).value

            If rowDateTime < nextStart Then
                ' Find the row ID that is the closest to nextStart
                rowStep = j
            End If

            ' Check if the date and time falls within the window
            If rowDateTime < winStart Then
                GoTo ContinueLoop
            End If

            If rowDateTime >= winEnd Then
                ' Exit the loop if the date and time is beyond the current window
                Exit For
            End If

            ' Check If the type matched
            If chkSomeAcc = True Then
                If strSomeAcc <> ws.Cells(j, ColShSimpleTSInOutAcc).value Then
                    GoTo ContinueLoop
                Else
                    ws.Cells(j, ColShSimpleTSInOutAcc).Interior.color = ColorBlue
                    ws.Cells(j, ColShSimpleTSInOutAcc).Font.color = ColorWhite
                End If
            End If

            Dim strChan As String
            strChan = ws.Cells(j, ColShSimpleTempTSCh).value

            ' 過濾交易通路與交易類型
            ' 目前可以選擇的是 交易通路: 臨櫃分行、ATM、自動化通路; 交易類型: 存入、提出、現金交易
            ' 以下是通路
            If IsTargetChannels(useBranch, useATM, useAuto, strChan) = False Then
                GoTo ContinueLoop
            End If

            'FIXME: check data type
            amtBalance = CDbl(ws.Cells(j, ColShSimpleBalance).value)
            amtWithdraw = CDbl(ws.Cells(j, ColShSimpleTSOut).value)
            amtDeposit = CDbl(ws.Cells(j, ColShSimpleTSIn).value)

            If Not (((typeDeposit = True) And (amtDeposit > 0)) Or _
               ((typeWithdraw = True) And (amtWithdraw > 0))) Then
                GoTo ContinueLoop
            End If

            If (typeCash = True) And (ws.Cells(j, ColShSimpleTSSummary).value <> ColTSSummaryVal01) Then
                GoTo ContinueLoop
            End If

            amtTxn = amtDeposit + amtWithdraw
            sumTxn = sumTxn + amtTxn
            If (chkSingleMax = True) And (amtTxn > amtSingleMax) Then
                GoTo ContinueLoop
            End If

            If (chkSingleMin = True) And (amtTxn < amtSingleMin) Then
                GoTo ContinueLoop
            End If

            accNo = ws.Cells(j, ColShSimpleTSInOutAcc)
            If accNo = "" Then
                accNo = strDefaultAccount
            End If

            If amtDeposit > 0 Then
                If bUnique = False Then
                    sumDeposit = sumDeposit + amtDeposit
                    ws.Cells(j, ColShSimpleTSSummary).AddComment UiCondiMatched
                Else
                    If dictDeposit.Exists(accNo) Then
                        If amtDeposit > dictDeposit(accNo) Then
                            dictDeposit(accNo) = amtDeposit
                            dictMarkDep(accNo) = j
                        End If
                    Else
                        dictDeposit.Add accNo, amtDeposit
                        dictMarkDep.Add accNo, j
                    End If

                End If
            End If

            If amtWithdraw > 0 Then
                If bUnique = False Then
                    sumWithdraw = sumWithdraw + amtWithdraw
                    ws.Cells(j, ColShSimpleTSSummary).AddComment UiCondiMatched
                Else
                    If dictWithdraw.Exists(accNo) Then
                        If amtDeposit > dictWithdraw(accNo) Then
                            dictWithdraw(accNo) = amtWithdraw
                            dictMarkWith(accNo) = j
                        End If
                    Else
                        dictWithdraw.Add accNo, amtWithdraw
                        dictMarkWith.Add accNo, j
                    End If
                End If
            End If

            amtBalance = ws.Cells(j, ColShSimpleBalance).value
            Utils.CountOne cntTxn
            Utils.CountOne cntHit  ' Check if the accumulated amount exceeds the threshold

            If hitTop = 0 Then
                hitTop = j
            End If
            hitBottom = j

            ' Do something (e.g., highlight the row or display a message)
            If (chkSingleMin = True) Or (chkSingleMax = True) Then
                If amtDeposit > 0 Then
                    ws.Cells(j, ColShSimpleTSIn).Font.color = ColorWhite
                    ws.Cells(j, ColShSimpleTSIn).Interior.color = ColorBlue
                End If
                If amtWithdraw > 0 Then
                    ws.Cells(j, ColShSimpleTSOut).Font.color = ColorWhite
                    ws.Cells(j, ColShSimpleTSOut).Interior.color = ColorBlue
                End If
            End If
ContinueLoop:
        Next j

        If cntHit < minOccur Then   ' If matched, frame it
            GoTo EndOfMainLoop
        End If

        If (chkSingleMin = False) And (chkSingleMax = False) And (chkAccMax = False) And _
           (chkDiff = False) And (chkBalanceMax = False) Then
            ' 個數超過
            DrawRedThickBorder Range("B" & CStr(i) & ":" & ColShSimpleLastCol & CStr(j - 1))
        End If

        If bUnique = True Then
            sumDeposit = Utils.SumDictValues(dictDeposit)
            sumWithdraw = Utils.SumDictValues(dictWithdraw)

            Dim item As Variant
            For Each item In dictMarkDep.items
                If Not (ws.Cells(item, ColShSimpleTSSummary).comment Is Nothing) Then
                    ws.Cells(item, ColShSimpleTSSummary).comment.Delete
                End If
                ws.Cells(item, ColShSimpleTSSummary).AddComment UiCondiMatched
            Next item

            For Each item In dictMarkWith.items
                If Not (ws.Cells(item, ColShSimpleTSSummary).comment Is Nothing) Then
                    ws.Cells(item, ColShSimpleTSSummary).comment.Delete
                End If
                ws.Cells(item, ColShSimpleTSSummary).AddComment UiCondiMatched
            Next item
        End If

        ' Check condition 取交集
        Dim bLowBalance As Boolean
        Dim bExceedAcc  As Boolean
        Dim bSmallDiff  As Boolean

        ' Check balance amtBalance
        If (chkBalanceMax = True) And (amtBalance < amtBalanceMax) Then
            ws.Cells(j - 1, ColShSimpleBalance).Font.color = ColorWhite
            ws.Cells(j - 1, ColShSimpleBalance).Interior.color = ColorBlue
            bLowBalance = True
        End If

        If chkAccMax = True Then
            If (typeDeposit = True) And (typeWithdraw = True) Then
                If sumTxn > amtAccMax Then
                    bExceedAcc = True
                End If
            ElseIf typeDeposit = True Then
                If sumDeposit > amtAccMax Then
                    bExceedAcc = True
                End If
            ElseIf typeWithdraw = True Then
                If sumWithdraw > amtAccMax Then
                    bExceedAcc = True
                End If
            End If
        End If

        If chkDiff = True Then
            If (sumWithdraw >= 0) And (sumDeposit >= 0) Then
                If Abs(sumWithdraw - sumDeposit) < amtDiff Then
                    bSmallDiff = True
                End If
            End If
        End If

        ' Draw borders
        If (chkAccMax = True) Or (chkDiff = True) Or (chkBalanceMax = True) Then
            If Not (chkAccMax Xor bExceedAcc) Then
                If Not (chkDiff Xor bSmallDiff) Then
                    If Not (chkBalanceMax Xor bLowBalance) Then
                        DrawRedThickBorder Range("B" & CStr(i) & ":" & ColShSimpleLastCol & CStr(j - 1))
                    End If
                End If
            End If
        End If
EndOfMainLoop:
        If rowStep <> 0 Then
            i = rowStep + 1        ' The first entry that exceeds nextStart
        End If
    Wend ' End While

    ' reset?
    Set dictDeposit = Nothing
    Set dictWithdraw = Nothing
End Sub

Private Sub ResetPage(ByVal ws As Worksheet)
    Dim cell As Range

    ' 重置字的顏色
    For Each cell In ws.UsedRange
        If Not IsEmpty(cell.value) Then
            cell.Font.color = ColorBlack
        End If
    Next cell

    ws.Rows(RowHeaderBegin).Font.color = ColorWhite

    ' 重置每行背景色
    Dim lastRow     As Long
    Dim i           As Long
    lastRow = ws.Cells(ws.Rows.Count, ColShSimpleTSTime).End(xlUp).row

    For i = RowDataBegin To lastRow
        ' set background color
        If (i - RowDataBegin) Mod 2 = 0 Then
            ws.Range(ws.Cells(i, "B"), ws.Cells(i, "K")).Interior.color = ColorYellow
        Else
            ws.Range(ws.Cells(i, "B"), ws.Cells(i, "K")).Interior.color = ColorWhite
        End If
    Next i

    ' Draw boards
    Dim TargetRange As Range
    Set TargetRange = ws.Range("A9:" & ColShSimpleLastCol & CStr(lastRow))
    With TargetRange.Borders
        .LineStyle = xlContinuous
        .color = ColorGreen2
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub

Private Sub DrawRedThickBorder(rng As Range)
    Dim color       As Long
    Dim line        As Integer

    Utils.CountOne gCntFrames

    If (gCntFrames Mod 3) = 0 Then
        color = ColorRed
        line = xlContinuous
    ElseIf (gCntFrames Mod 3) = 1 Then
        color = ColorBlue
        line = xlDash
    Else
        color = ColorGreen
        line = xlContinuous
    End If

    ' Apply red, thick border to the range
    With rng.Borders(xlEdgeTop)
        .LineStyle = line
        .color = color
        .Weight = xlMedium        ' Thick border
    End With

    With rng.Borders(xlEdgeBottom)
        .LineStyle = line
        .color = color
        .Weight = xlMedium        ' Thick border
    End With

    With rng.Borders(xlEdgeLeft)
        .LineStyle = line
        .color = color
        .Weight = xlMedium        ' Thick border
    End With

    With rng.Borders(xlEdgeRight)
        .LineStyle = line
        .color = color
        .Weight = xlMedium        ' Thick border
    End With
End Sub

Private Sub Worksheet_Activate()
    gLabelCur = Range(cbLabel).value

    If UpdateSimplePage = False Then
        UpdateSimplePage = True ' lock
        Dim blScreenUpdating As Boolean
        blScreenUpdating = Application.ScreenUpdating
        If blScreenUpdating = True Then
            ' Fix flash issue when switching to this page at the beginning
            Application.ScreenUpdating = False
        End If

        Refresh_Click_Int False

        Application.ScreenUpdating = blScreenUpdating ' restore value

        UpdateSimplePage = False ' Unlock
    End If

    Dim ws   As Worksheet
    Dim lastRow As Long
    Set ws = ThisWorkbook.Sheets(SheetNameSimple)
    lastRow = ExcelUtils.GetLastRowNumber(ws)
    If lastRow >= RowDataBegin Then
        ExcelUtils.ApplySheetSimpleStyle
    End If
End Sub
