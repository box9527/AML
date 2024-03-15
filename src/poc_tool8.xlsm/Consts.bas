Attribute VB_Name = "Consts"
Option Explicit

' If you cannot see the following Chinese words,
' please change your encoding from UTF-8 to Big5-HKSCS.
' 如果你看不到這段文字，請將你的 encoding 從 UTF-8 改為 Big5-HKSCS。

' 特別的定義值，用來 initial 一個空array 時給定的初始值。
Public Const EmptyArrayValue As Long = -1

' Derivative watch-listed account
Public Const FileBadAcc      As String = "警示戶.xlsx"
Public Const SheetNameBadAcc As String = "警示戶"

' Virtual account
Public Const FileVirtualAcc      As String = "虛擬帳戶.xlsx"
Public Const SheetNameVirtualAcc As String = "照會資料"

' Coloring cells
' Note: The order is BGR if you use hex
Public Const ColorLightGray As Long = &HF0F0F0 'RGB(240, 240, 240)
Public Const ColorRed       As Long = &H101FF  'RGB(255, 0, 0)
Public Const ColorYellow    As Long = &HDDFFFF '&HC4F7F6 ' RGB(246, 247, 196)
Public Const ColorWhite     As Long = &HFFFFFF 'RGB(255, 255, 255)
Public Const ColorPink      As Long = &H9912E1 '&HA252FF
Public Const ColorBlue      As Long = &HAD5236
Public Const ColorOrange    As Long = &HB60B0
Public Const ColorBlack     As Long = &H0
Public Const ColorGreen     As Long = &H7C7C00 'RGB(0, 124, 124)
Public Const ColorYellow2   As Long = &H86FFFE

' 主要的 Sheet 名字
Public Const SheetNameOrginal      As String = "1原始資料"          ' 目前跑程式加入
Public Const SheetNameMain         As String = "2.1主頁面"          ' 常駐出現
Public Const SheetNameInputData    As String = "2.2清整後資料"      ' 常駐出現
Public Const SheetNameSimple       As String = "3.1交易明細"        ' 常駐出現
Public Const SheetNameMoney        As String = "3.2金流與交易對手"   ' 常駐出現
Public Const SheetNameBranch       As String = "分行清單"           ' 隱藏
Public Const SheetNameIntermediate As String = "暫存區"             ' 隱藏
Public Const SheetNameLabel        As String = "自訂標示設定"        ' 隱藏

' 原始資料從第幾行開始才是真的資料
Public Const RowDataBegin       As Integer = 9

' Sheet "清整後資料" 裡的欄位位置
Public Const ColTSDate         As Integer = 1        ' 交易日期
Public Const ColAccDate        As Integer = 2        ' 帳務日期
Public Const ColTSCode         As Integer = 3        ' 交易代碼
Public Const ColTSTime         As Integer = 4        ' 交易時間
Public Const ColBranchID       As Integer = 5        ' 交易分行
Public Const ColTSTeller       As Integer = 6        ' 交易櫃員
Public Const ColSummary        As Integer = 7        ' 摘要
Public Const ColAmount         As Integer = 8        ' 金額
Public Const ColBalance        As Integer = 9        ' 餘額
Public Const ColAccount        As Integer = 10       ' 轉出入帳戶
Public Const ColMemberID       As Integer = 11       ' 合作機構/會員編號
Public Const ColSerialCode     As Integer = 12       ' 金資序號
Public Const ColChannel        As Integer = 13       ' 註記
Public Const ColNote           As Integer = 14       ' 備註
Public Const ColTSMonth        As Integer = 15       ' 交易月份
Public Const ColTSSummary      As Integer = 16       ' 交易摘要
Public Const ColAmountTransfer As Integer = 17       ' 轉出金額
Public Const ColAmountDeposit  As Integer = 18       ' 存入金額
Public Const ColBankCode       As Integer = 19       ' 銀行代碼
Public Const ColTSType         As Integer = 20       ' TranType
Public Const ColATMLoc         As Integer = 21       ' ATM地點
Public Const ColATMCity        As Integer = 22       ' ATM縣市
Public Const ColATMArea        As Integer = 23       ' ATM區域
Public Const ColBranchName     As Integer = 24       ' 分行名
Public Const ColBranchCity     As Integer = 25       ' 分行縣市
Public Const ColBranchArea     As Integer = 26       ' 分行區域
Public Const ColTSLoc          As Integer = 27       ' 交易地點
Public Const ColTSChannel      As Integer = 28       ' 交易通路
Public Const ColTSOClock       As Integer = 29       ' 簡易交易時間

'假設原始資料最多的 Column 數目就是最多 50 欄
Public Const MaxSrcCol         As Integer = 50

' 這兩個 交易分行代碼 基本上根據 BU 的說法
' 是指 "自動化交易" 這類的分行代碼都會是0880 對 BU 來說非必要的資訊
' 所以特別獨立出來，方便後面統一過濾
Public Const SelfServiceID   As String = "0880"      ' 您幫幫您
Public Const SelfServiceID_2 As String = "880"       ' 您幫幫您

' =========================================================
' Sheet 3.1 交易明細 所使用的常數
' 顯示警告原因的直欄號碼
Public Const ColAlertReason As String = "K"

' 各式警告原因
Public Const ReasonCloseTo50w     As String = "金額接近 50 萬"
Public Const ReasonSmallTSAmount  As String = "小額轉帳"
Public Const ReasonTSFast         As String = "快速進出"
Public Const ReasonTSInMorning    As String = "凌晨交易"
Public Const ReasonTSLargeAmount  As String = "大額進出"
Public Const ReasonDWLAlert       As String = "往來警示帳戶"
' =========================================================

' =========================================================
' Sheet 暫存區 所使用的常數
' 這裡的常數，是為了要來畫出 2.1 主頁面 各種圖，所要暫時紀錄的 "欄位" "值" 是放暫存區哪裡
' 例如:
' RowTask1TotalCountXAxis & RowTask1TotalCountYAxis 
' 就是把 分月交易總次額 的橫軸 (日期) 縱軸 (次數) 放到暫存區的 row 4 & row 5
' 切去 暫存區 就可以理解

' 任務1, 分月交易總次數
Public Const RowTask1TotalCountXAxis As Integer = 4             ' 紀載在第 4 列
Public Const RowTask1TotalCountYAxis As Integer = 5             ' 紀載在第 5 列

' 任務2, 分月交易總金額
' 任務3, 可疑大額提領
Public Const RowTask23TotalAmountXAxis As Integer = 7           ' 紀載在第 7 列
Public Const RowTask23TotalAmountYAxis As Integer = 8           ' 紀載在第 8 列

' 任務4, 當日現金提領_臨櫃 (臨櫃提款)
Public Const RowTask4WithdrawOverCounterXAxis As Integer = 10   ' 紀載在第 10 列
Public Const RowTask4WithdrawOverCounterYAxis As Integer = 11   ' 紀載在第 11 列

' 任務5, 當日現金提領_臨櫃 (提款總額)
Public Const RowTask5WithdrawSummaryXAxis As Integer = 13       ' 紀載在第 13 列
Public Const RowTask5WithdrawSummaryYAxis As Integer = 14       ' 紀載在第 14 列

' 任務5, 當日現金提領_臨櫃 (提款分行)
Public Const RowTask5WithdrawBranchXAxis As Integer = 16        ' 紀載在第 16 列
Public Const RowTask5WithdrawBranchYAxis As Integer = 17        ' 紀載在第 17 列

' 任務5, 去過的分行統計
Public Const RowTask5VisitedBranchXAxis As Integer = 19         ' 紀載在第 19 列
Public Const RowTask5VisitedBranchYAxis As Integer = 20         ' 紀載在第 20 列

' 任務6, 當日現金提領/存入_ATM (快進快出)
Public Const RowTask6QuickXAxis As Integer = 42                 ' 紀載在第 42 列
Public Const RowTask6QuickYAxis As Integer = 43                 ' 紀載在第 43 列

' 任務7, 當日現金提領/存入_ATM (出入帳金額比)
Public Const RowTask7DiffXAxis    As Integer = 36               ' 紀載在第 36 列
Public Const RowTask7DiffTransfer As Integer = 37               ' 紀載在第 37 列 ' 轉帳金額Y軸的值
Public Const RowTask7DiffDeposit  As Integer = 38               ' 紀載在第 38 列 ' 存現金額Y軸的值
Public Const RowTask7DiffRatio    As Integer = 39               ' 紀載在第 39 列 ' 出入帳金額比例Y軸的值

' 任務9, 當日現金提領/存入_ATM
Public Const RowTask9ATMDepositXAxis As Integer = 22            ' 紀載在第 22 列
Public Const RowTask9ATMDepositYAxis As Integer = 23            ' 紀載在第 23 列
Public Const RowTask9ATMDepositLocation As Integer = 24         ' 紀載在第 24 列 ' 存現地點Y軸的值
Public Const RowTask9ATMDepositLocationDetail As Integer = 25   ' 紀載在第 25 列 ' 存現地點詳細資料Y軸的值

' 任務11, 當日存款餘額
Public Const RowTask11BalanceXAxis As Integer = 27              ' 紀載在第 27 列
Public Const RowTask11BalanceYAxis As Integer = 28              ' 紀載在第 28 列

' 任務12, 當日轉帳金額
Public Const RowTask12SmallTransferXAxis     As Integer = 30    ' 紀載在第 30 列
Public Const RowTask12SmallTransferOneDollar As Integer = 31    ' 紀載在第 31 列 ' 一元交易Y軸的值
Public Const RowTask12SmallTransferHundred   As Integer = 32    ' 紀載在第 32 列 ' 小於百元交易Y軸的值
Public Const RowTask12SmallTransferMany      As Integer = 33    ' 紀載在第 33 列 ' 跨行交易Y軸的值

' Note: this one has to be the last as the item count is not fixed
' 3.2 交易對手 紀錄的值所起始的列號，請注意，如果上面 2.1 的列記錄快牴觸到的時候，這個值是可以往下調整的。
Public Const RowTopCounterparty As Integer = 60                 ' 紀載在第 60 列 ' 從這列以下，是3.2的值
' =========================================================