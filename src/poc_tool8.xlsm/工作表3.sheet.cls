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
Const colDay        As String = "A"
Const colTxnType    As String = "B"
Const colTime       As String = "C"
Const colWithdraw   As String = "D"
Const colDeposit    As String = "E"
Const colBalance    As String = "F"
Const colAccount    As String = "G"
Const colReason     As String = "K"
Const colChannel     As String = "Y"
Const colBankID     As String = "Z"

Const lastCol       As String = "K"

Const chkEnableBalance As String = "P1"
Const chkEnableTxnTime As String = "P2"
Const chkEnableLarge As String = "P3"
Const chkEnableNearLarge As String = "P4"
Const chkUnique As String = "P5"
Const txtSomeAcc As String = "Q1"

Const cbWin         As String = "R1"
Const cbOccur       As String = "R2"
Const cbOppo        As String = "R3"
Const chkBranch As String = "S1"
Const chkATM As String = "S2"
Const chkAuto As String = "S3"
Const chkDeposit As String = "T1"
Const chkWithdraw As String = "T2"
Const chkCash As String = "T3"
Const txtAmtMin As String = "U1"
Const txtAmtMax As String = "U2"
Const txtAmtDf As String = "U3"
Const txtAmtLessThan As String = "U4"
Const txtAmtTotal As String = "U5"
Const cbLabel As String = "V1"
Const strDisableSearch As String = "不使用"

Dim cntFrames       As Long

' Row index in config sheet
Const RowInterval       As String = "B"
Const RowTimes As String = "C"
Const RowOppo As String = "D"
Const RowUnique As String = "E"
Const RowSomeAcc As String = "F"
Const RowBranch As String = "G"
Const RowATM As String = "H"
Const RowAuto As String = "I"
Const RowDeposit As String = "J"
Const RowWithdraw As String = "K"
Const RowCash As String = "L"
Const RowAmtMin As String = "M"
Const RowAmtMax As String = "N"
Const RowAmtDf As String = "O"
Const RowAmtLessThan As String = "P"
Const RowAmtTotal As String = "Q"
Const RowDesc As String = "R"

Dim gLabelCur        As String


Private Sub cbPattern_Change()

    Debug.Print "Hit change" & Range(cbLabel)

    RefreshSearchSetting Range(cbLabel)
End Sub

Private Sub RefreshSearchSetting(strLabel)
    Dim wsCfg As Worksheet

    ' Find the row
    Dim rowCfg As Long
    Dim lastRow As Long
    Dim i As Long

    Set wsCfg = Worksheets(SheetNameLabel)

    If strLabel = strDisableSearch Then
        gLabelCur = Range(cbLabel).Value
        Exit Sub
    End If

    rowCfg = 1
    lastRow = wsCfg.Cells(wsCfg.Rows.Count, 1).End(xlUp).row

    For i = 2 To lastRow
        If wsCfg.Cells(i, 1).Value = strLabel Then
            rowCfg = i
            Exit For
        End If
    Next i

    If rowCfg = 1 Then
        MsgBox "無法在自訂標示設定頁找到對應設定:" & strLabel
        Range(cbLabel).Value = gLabelCur
        'FIXME: should update pattern description here
        Exit Sub
    End If

    ' Fill row value to UI position
    Range(cbWin).Value = wsCfg.Cells(rowCfg, RowInterval)
    Range(cbOccur).Value = wsCfg.Cells(rowCfg, RowTimes)
    Range(cbOppo).Value = wsCfg.Cells(rowCfg, RowOppo)
    Range(chkUnique).Value = wsCfg.Cells(rowCfg, RowUnique)
    Range(txtSomeAcc).Value = wsCfg.Cells(rowCfg, RowSomeAcc)
    Range(chkBranch).Value = wsCfg.Cells(rowCfg, RowBranch)
    Range(chkATM).Value = wsCfg.Cells(rowCfg, RowATM)
    Range(chkAuto).Value = wsCfg.Cells(rowCfg, RowAuto)
    Range(chkDeposit).Value = wsCfg.Cells(rowCfg, RowDeposit)
    Range(chkWithdraw).Value = wsCfg.Cells(rowCfg, RowWithdraw)
    Range(chkCash).Value = wsCfg.Cells(rowCfg, RowCash)
    Range(txtAmtMin).Value = wsCfg.Cells(rowCfg, RowAmtMin)
    Range(txtAmtMax).Value = wsCfg.Cells(rowCfg, RowAmtMax)
    Range(txtAmtDf).Value = wsCfg.Cells(rowCfg, RowAmtDf)
    Range(txtAmtLessThan).Value = wsCfg.Cells(rowCfg, RowAmtLessThan)
    Range(txtAmtTotal).Value = wsCfg.Cells(rowCfg, RowAmtTotal)

    Worksheets(SheetNameSimple).OLEObjects("ptnDesc").Object.caption = wsCfg.Cells(rowCfg, RowDesc).Value

    gLabelCur = Range(cbLabel).Value
End Sub

Private Sub Refresh_Click()
    Refresh_Click_Int True
End Sub

Private Sub Refresh_Click_Int(blShowOk As Boolean)
    Dim ws          As Worksheet
    Dim enableSearch As Boolean

    If Range(cbLabel).Value = strDisableSearch Then
        enableSearch = False
    Else
        enableSearch = True
    End If

    ExtraData.InitDicts

    Set ws = ThisWorkbook.Sheets(SheetNameSimple)

    cntFrames = 0
    Application.ScreenUpdating = False        ' for performance
    ' Reset all cells
    ResetPage ws

    CheckOneByOne ws

    If enableSearch = True Then

        TraverseRowsByTimeWindow ws


    End If
    Application.ScreenUpdating = True
    If blShowOk = True Then
        MsgBox "完成"
    End If

End Sub

Sub CheckOneByOne(ByVal ws As Worksheet)
    Dim i           As Long
    Dim lastRow     As Long
    Dim cellTime    As Double
    Dim strTmp      As String
    Dim strAccount  As String
    Dim enableBalance As Boolean
    Dim enableTxnTime As Boolean
    Dim enableLarge As Boolean
    Dim enableNearLarge As Boolean

    enableBalance = Range(chkEnableBalance).Value
    enableTxnTime = Range(chkEnableTxnTime).Value
    enableLarge = Range(chkEnableLarge).Value
    enableNearLarge = Range(chkEnableNearLarge).Value

    lastRow = ws.Cells(ws.Rows.Count, colTime).End(xlUp).row

    For i = RowDataBegin To lastRow
        ' Check balance
        If enableBalance = True Then
            If VarType(ws.Cells(i, colBalance).Value) = vbDouble Then
                If ws.Cells(i, colBalance).Value < 1000 Then
                    ws.Cells(i, colBalance).Font.color = ColorRed
                End If
            End If
        End If
        ' Check amount
        If VarType(ws.Cells(i, colDeposit).Value) = vbDouble And Not IsEmpty(ws.Cells(i, colDeposit)) Then
            If ws.Cells(i, colDeposit).Value < 500000 And ws.Cells(i, colDeposit).Value >= 400000 Then
                ws.Cells(i, colDeposit).Font.color = ColorRed
            End If

            If ws.Cells(i, colDeposit).Value >= 500000 Then
                ws.Cells(i, colDeposit).Font.color = ColorRed
            End If
        End If

        If VarType(ws.Cells(i, colWithdraw).Value) = vbDouble And Not IsEmpty(ws.Cells(i, colWithdraw)) Then
            If enableNearLarge = True Then
                If ws.Cells(i, colWithdraw).Value < 500000 And ws.Cells(i, colWithdraw).Value >= 400000 Then
                    ws.Cells(i, colWithdraw).Font.color = ColorRed
                End If
            End If

            If enableLarge = True Then
                If ws.Cells(i, colWithdraw).Value >= 500000 Then
                    ws.Cells(i, colWithdraw).Font.color = ColorRed
                End If
            End If
        End If

        ' Check time
        If enableTxnTime = True Then
            If VarType(ws.Cells(i, colTime).Value) = vbDouble Then
                cellTime = ws.Cells(i, colTime).Value
                If cellTime >= TimeValue("0:00:00") And cellTime <= TimeValue("4:00:00") Then
                    ws.Cells(i, colTime).Font.color = ColorRed
                End If
            End If
        End If

        ' Check bad accounts
        If VarType(ws.Cells(i, colAccount).Value) = vbString Then
            strAccount = ws.Cells(i, colAccount).Value
            If Not IsEmpty(strAccount) Then
                If DictBlacklist.Exists(strAccount) Then
                    'Debug.Print " 警示戶　hit!!! " & strAccount
                    ws.Cells(i, colAccount).Font.color = ColorWhite
                    ws.Cells(i, colAccount).Interior.color = ColorRed
                    ws.Cells(i, colReason).Value = "警示戶"
                End If
            End If

            strTmp = ExtraData.CheckVAccount(ws.Cells(i, colBankID).Value, strAccount)
            If strTmp <> "" Then
                ws.Cells(i, colAccount).Font.color = ColorWhite
                ws.Cells(i, colAccount).Interior.color = ColorRed

                ws.Cells(i, colReason).Value = strTmp
            End If

        End If

    Next i

End Sub

Sub TraverseRowsByTimeWindow(ByVal ws As Worksheet)
    Dim lastRow     As Long
    Dim currentTime As Date
    Dim winStart    As Date
    Dim winEnd      As Date

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

    Dim useBranch   As Boolean
    Dim useATM      As Boolean
    Dim useAuto     As Boolean
    Dim typeDeposit As Boolean
    Dim typeWithdraw As Boolean
    Dim typeCash    As Boolean

    Dim bUnique As Boolean

    Dim amtSingleUppB As Double
    Dim amtSingleLowB As Double
    Dim amtDiff     As Double
    Dim amtBalanceMax As Double
    Dim amtAccMax   As Double

    Dim chkSomeAcc  As Boolean
    Dim strSomeAcc  As String

    Dim chkSingleUppB As Boolean
    Dim chkSingleLowB As Boolean
    Dim chkDiff     As Boolean
    Dim chkBalanceMax As Boolean
    Dim chkAccMax   As Boolean

    strWindow = Range(cbWin).Value
    strOccur = Range(cbOccur).Value
    strOpponent = Range(cbOppo).Value
    useBranch = Range(chkBranch).Value
    useATM = Range(chkATM).Value
    useAuto = Range(chkAuto).Value
    typeDeposit = Range(chkDeposit).Value
    typeWithdraw = Range(chkWithdraw).Value
    typeCash = Range(chkCash).Value

    bUnique = Range(chkUnique).Value


    chkSomeAcc = False
    strSomeAcc = ""
    If Trim(Range(txtSomeAcc)) <> "" And Range(cbOppo) = "指定帳戶" Then
        strSomeAcc = Trim(Range(txtSomeAcc))
        chkSomeAcc = True
    End If

    chkSingleLowB = False
    chkSingleUppB = False
    chkDiff = False
    chkBalanceMax = False
    chkAccMax = False
    If Trim(Range(txtAmtMin)) <> "" Then
        chkSingleLowB = True
        amtSingleLowB = Range(txtAmtMin)
    End If

    If Trim(Range(txtAmtMax)) <> "" Then
        chkSingleUppB = True
        amtSingleUppB = Range(txtAmtMax)
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
        MsgBox "當差值數字不等於0時, [存入]與[提出]必須被勾選"
        Exit Sub
    End If

    If strWindow = "每日" Then
        timeUnit = "日"
        winLength = 1        ' Set the time window in hours
    ElseIf strWindow = "每3日" Then
        timeUnit = "日"
        winLength = 3        ' Set the time window in hours
    ElseIf strWindow = "每5日" Then
        timeUnit = "日"
        winLength = 5        ' Set the time window in hours
    ElseIf strWindow = "每月" Then
        timeUnit = "月"
        winLength = 1        ' Set the time window in hours
    Else
        MsgBox "unknown window setting!" & strWindow
        Exit Sub
    End If

    If strOccur = "至少1次" Then
        minOccur = 1
    ElseIf strOccur = "至少3次" Then
        minOccur = 3
    ElseIf strOccur = "至少5次" Then
        minOccur = 5
    Else
        MsgBox "unknown occurrence setting!" & strOccur
        Exit Sub
    End If





    ' Find the last row with data
    lastRow = ws.Cells(ws.Rows.Count, colTime).End(xlUp).row

    Dim dictDeposit As Object
    Dim dictWithdraw As Object
    Dim dictMarkDep As Object
    Dim dictMarkWith As Object

    Set dictDeposit = CreateObject("Scripting.Dictionary")
    Set dictWithdraw = CreateObject("Scripting.Dictionary")
    Set dictMarkDep = CreateObject("Scripting.Dictionary")
    Set dictMarkWith = CreateObject("Scripting.Dictionary")



    ' Loop through the rows
    nextStart = ws.Cells(RowDataBegin, colDay).MergeArea.Cells(1, 1).Value + ws.Cells(RowDataBegin, colTime).Value
    nextRowStart = RowDataBegin
    i = RowDataBegin
    While i <= lastRow
        ' Get the current date and time; Fetch the real row
        currentTime = ws.Cells(i, colDay).MergeArea.Cells(1, 1).Value + ws.Cells(i, colTime).Value
        ' Calculate the window boundaries
        If currentTime >= nextEnd Then
            winStart = currentTime
        Else
            winStart = nextStart
        End If
        If timeUnit = "時" Then
            winEnd = DateAdd("h", winLength, winStart)
            nextStart = DateAdd("h", 1, winStart)
            nextEnd = nextStart + DateAdd("h", winLength, nextStart)

        ElseIf timeUnit = "日" Then
            winEnd = DateSerial(year(winStart), month(winStart), Day(winStart) + winLength) + _
                     TimeSerial(0, 0, 0)
            nextStart = DateSerial(year(winStart), month(winStart), Day(winStart) + 1) + _
                     TimeSerial(0, 0, 0)
            nextEnd = DateSerial(year(nextStart), month(nextStart), Day(nextStart) + winLength) + _
                     TimeSerial(0, 0, 0)

        ElseIf timeUnit = "月" Then
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

        Dim amtDeposit As Double
        Dim sumDeposit As Double
        Dim amtWithdraw As Double
        Dim sumWithdraw As Double
        Dim amtBalance As Double
        Dim amtTxn  As Double
        Dim sumTxn  As Double
        Dim hitTop  As Long
        Dim hitBottom As Long

        hitTop = 0
        hitBottom = 0

        cntTxn = 0
        cntHit = 0
        amtDeposit = 0
        sumDeposit = 0
        amtWithdraw = 0
        sumWithdraw = 0
        amtTxn = 0
        sumTxn = 0

        amtBalance = 0

        dictDeposit.RemoveAll
        dictWithdraw.RemoveAll
        dictMarkDep.RemoveAll
        dictMarkWith.RemoveAll


        ' Traverse rows within the current time window
        'Debug.Print "every " & CStr(winLength) & " " & timeUnit
        'Debug.Print "start " & winStart
        'Debug.Print "end " & winEnd

        Dim rowStep As Long
        Dim accNo As String
        rowStep = 0
        For j = i To lastRow

            If Not (ws.Cells(j, colTxnType).comment Is Nothing) Then
                ws.Cells(j, colTxnType).comment.Delete
            End If
            If Not (ws.Cells(j, colTxnType).comment Is Nothing) Then
                ws.Cells(j, colTxnType).comment.Delete
            End If

            ' Get the date and time of the current row
            Dim rowDateTime As Date
            rowDateTime = ws.Cells(j, colDay).MergeArea.Cells(1, 1).Value + ws.Cells(j, colTime).Value

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

            ' Debug.Print rowDateTime

            ' Check If the type matched
            If chkSomeAcc = True Then
                If strSomeAcc <> ws.Cells(j, colAccount).Value Then
                    GoTo ContinueLoop
                Else
                    ws.Cells(j, colAccount).Interior.color = ColorBlue
                    ws.Cells(j, colAccount).Font.color = ColorWhite
                End If
            End If
            Dim strChan As String

            strChan = ws.Cells(j, colChannel).Value

            If Not (((useBranch = True) And ((strChan = "臨櫃分行") Or (strChan = "電匯"))) Or _
                   ((useATM = True) And ((strChan = "ATM自動化設備") Or (strChan = "跨行機台"))) Or _
                   ((useAuto = True) And ((strChan = "行動網") Or (strChan = "網銀")))) Then

                GoTo ContinueLoop
            End If

            'FIXME: check data type
            amtBalance = CDbl(ws.Cells(j, colBalance).Value)
            amtWithdraw = CDbl(ws.Cells(j, colWithdraw).Value)
            amtDeposit = CDbl(ws.Cells(j, colDeposit).Value)

            If Not (((typeDeposit = True) And (amtDeposit > 0)) Or _
               ((typeWithdraw = True) And (amtWithdraw > 0))) _
               Then
                GoTo ContinueLoop
            End If

            If (typeCash = True) And (ws.Cells(j, colTxnType).Value <> "01現金交易") Then
                GoTo ContinueLoop
            End If

            amtTxn = amtDeposit + amtWithdraw
            sumTxn = sumTxn + amtTxn
            If (chkSingleUppB = True) And (amtTxn > amtSingleUppB) Then
                GoTo ContinueLoop
            End If

            If (chkSingleLowB = True) And (amtTxn < amtSingleLowB) Then
                GoTo ContinueLoop
            End If

            accNo = ws.Cells(j, colAccount)
            If accNo = "" Then
                accNo = "self"
            End If

            If amtDeposit > 0 Then
                If bUnique = False Then
                    sumDeposit = sumDeposit + amtDeposit
                    ws.Cells(j, colTxnType).AddComment "符合條件"
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
                    ws.Cells(j, colTxnType).AddComment "符合條件"
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

            amtBalance = ws.Cells(j, colBalance).Value

            cntTxn = cntTxn + 1


            ' Check if the accumulated amount exceeds the threshold
            cntHit = cntHit + 1
            If hitTop = 0 Then
                hitTop = j
            End If
            hitBottom = j

            ' Do something (e.g., highlight the row or display a message)

            If (chkSingleLowB = True) Or (chkSingleUppB = True) Then
                If amtDeposit > 0 Then
                    ws.Cells(j, colDeposit).Font.color = ColorWhite
                    ws.Cells(j, colDeposit).Interior.color = ColorBlue
                End If
                If amtWithdraw > 0 Then
                    ws.Cells(j, colWithdraw).Font.color = ColorWhite
                    ws.Cells(j, colWithdraw).Interior.color = ColorBlue
                End If
            End If
ContinueLoop:
        Next j


        If cntHit < minOccur Then        ' If matched, frame it
            GoTo EndOfMainLoop
        Else
            If (chkSingleLowB = False) And (chkSingleUppB = False) And (chkAccMax = False) And (chkDiff = False) And (chkBalanceMax = False) Then
                DrawRedThickBorder Range("B" & CStr(i) & ":" & lastCol & CStr(j - 1))
            End If
        End If

        If bUnique = True Then
            sumDeposit = SumDictValues(dictDeposit)

            sumWithdraw = SumDictValues(dictWithdraw)

            Dim item
            For Each item In dictMarkDep.items
                If Not (ws.Cells(item, colTxnType).comment Is Nothing) Then
                    ws.Cells(item, colTxnType).comment.Delete
                End If
                ws.Cells(item, colTxnType).AddComment "符合條件"
            Next item
            For Each item In dictMarkWith.items
                If Not (ws.Cells(item, colTxnType).comment Is Nothing) Then
                    ws.Cells(item, colTxnType).comment.Delete
                End If
                ws.Cells(item, colTxnType).AddComment "符合條件"
            Next item

        End If

        ' Check condition
        Dim bExceedAcc  As Boolean
        bExceedAcc = False
        If chkAccMax = True Then

            If typeDeposit = True And typeWithdraw = True Then
                If (sumTxn) > amtAccMax Then
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
            If (bExceedAcc = True) And (chkDiff = False) Then
                DrawRedThickBorder Range("B" & CStr(i) & ":" & lastCol & CStr(j - 1))
            End If
        End If

        If chkDiff = True Then
            If (sumWithdraw > 0) And (sumDeposit > 0) Then
                If Abs(sumWithdraw - sumDeposit) < amtDiff Then
                    If (chkAccMax = False) Or ((chkAccMax = False) And (bExceedAcc = True)) Then
                        DrawRedThickBorder Range("B" & CStr(i) & ":" & lastCol & CStr(j - 1))
                    End If
                End If
            End If
        End If

        ' Check balance amtBalance
        If (chkBalanceMax = True) And (amtBalance < amtBalanceMax) Then
            ws.Cells(j - 1, colBalance).Font.color = ColorWhite
            ws.Cells(j - 1, colBalance).Interior.color = ColorBlue
            DrawRedThickBorder Range("B" & CStr(j - 1) & ":" & lastCol & CStr(j - 1))
        End If

EndOfMainLoop:
        If rowStep <> 0 Then
            i = rowStep + 1        ' The first entry that exceeds nextStart
        End If

    Wend
    Set dictDeposit = Nothing
    Set dictWithdraw = Nothing

End Sub

Function SumDictValues(dict As Object) As Double
    Dim key As Variant
    Dim total As Double
    total = 0
    ' Iterate through each item in the dictionary
    For Each key In dict.keys
        ' Add the value of the current item to the total
        total = total + dict(key)
    Next key
    ' Return the total sum
    SumDictValues = total
End Function

Sub ResetPage(ByVal ws As Worksheet)
    Dim cell        As Range

    For Each cell In ws.UsedRange
        If Not IsEmpty(cell.Value) Then
            cell.Font.color = ColorBlack
        End If
    Next cell
    ws.Rows(8).Font.color = ColorWhite

    Dim lastRow     As Long
    Dim i           As Long

    lastRow = ws.Cells(ws.Rows.Count, colTime).End(xlUp).row

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
    Set TargetRange = ws.Range("A9:" & lastCol & CStr(lastRow))
    With TargetRange.Borders
        .LineStyle = xlContinuous
        .color = RGB(153, 188, 133)        ' Green border color
        .TintAndShade = 0
        .Weight = xlThin
    End With


End Sub

'========

Sub DrawRedThickBorder(rng As Range)
    Dim color       As Long
    Dim line        As Integer
    cntFrames = cntFrames + 1

    If (cntFrames Mod 3) = 0 Then
        color = ColorRed
        line = xlContinuous
    ElseIf (cntFrames Mod 3) = 1 Then
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
    gLabelCur = Range(cbLabel).Value
    If UpdateSimplePage = False Then
        Refresh_Click_Int False
        UpdateSimplePage = True
    End If
End Sub
