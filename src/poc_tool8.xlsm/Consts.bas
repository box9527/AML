Attribute VB_Name = "Consts"
Option Explicit

' If you cannot see the following Chinese words,
' please change your encoding from UTF-8 to Big5-HKSCS.
' 如果你看不到這段文字，請將你的 encoding 從 UTF-8 改為 Big5-HKSCS。

' 特別的定義值，用來 initial 一個空array 時給定的初始值。
Public Const EmptyArrayValue As Long = -1

' Excel style
Public Const FontSize          As Integer = 12
Public Const FontName          As String = "微軟正黑體"
Public Const CharTitleFontSize As Integer = 18
Public Const CharAreaFontSize  As Integer = 14
Public Const AnalyBtnFontSize  As Integer = 24
Public Const DateFormat        As String = "yyyy/mm/dd"
Public Const TimeFormat        As String = "hh:mm:ss"
Public Const NumberFormat      As String = "_(* #,##0.00_);_(* (#,##0.00);_(* "" - ""??_);_(@_)"
Public Const MoneyFormat       As String = "$#,##0.00;-$#,##0.00"
Public Const GeneralFormat     As String = "General"
Public Const ForceStringFormat As String = "@"

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
Public Const ColorGreen2    As Long = &H99BC85 'RGB(153, 188, 133)
Public Const ColorYellow2   As Long = &H86FFFE

' 主要的 Sheet 名字
Public Const SheetNameOrginal      As String = "1原始資料"          ' 目前跑程式加入
Public Const SheetNameMain         As String = "2.1主頁面"          ' 常駐出現
Public Const SheetNameInputData    As String = "2.2清整後資料"      ' 常駐出現
Public Const SheetNameSimple       As String = "3.1交易明細"        ' 常駐出現
Public Const SheetNameMoney        As String = "3.2金流與交易對手"   ' 常駐出現
Public Const SheetNameBranch       As String = "分行清單"           ' 隱藏
Public Const SheetNameATM          As String = "ATM清單"           ' 隱藏
Public Const SheetNameIntermediate As String = "暫存區"             ' 隱藏
Public Const SheetNameLabel        As String = "自訂標示設定"        ' 隱藏

' 原始資料從第幾行開始才是真的資料
Public Const RowDataBegin As Long = 7 ' 20240403, 配合3RC從 9 改為 7
Public Const RowHeaderBegin As Long = (RowDataBegin - 1)

' =========================================================
' Sheet "1原始資料" 裡的欄位位置
Public Const ColShOrgTSDate     As Integer = 1       ' 交易日期
Public Const ColShOrgAccDate    As Integer = 2       ' 帳務日期
Public Const ColShOrgTSCode     As Integer = 3       ' 交易代碼
Public Const ColShOrgTSTime     As Integer = 4       ' 交易時間
Public Const ColShOrgBranchID   As Integer = 5       ' 交易分行
Public Const ColShOrgTSTeller   As Integer = 6       ' 交易櫃員
Public Const ColShOrgSummary    As Integer = 7       ' 摘要
Public Const ColShOrgAmtDraw    As Integer = 8       ' Out ' 20240403 配合3RC改為 支出
Public Const ColShOrgAmtDeposit As Integer = 9       ' In  ' 20240403 配合3RC改為 存入
Public Const ColShOrgBalance    As Integer = 10      ' 餘額
Public Const ColShOrgBankCat    As Integer = 11      ' 行庫別，20240403 配合3RC多出的Column
Public Const ColShOrgAccount    As Integer = 12      ' 轉出入帳號
Public Const ColShOrgMemberID   As Integer = 13      ' 合作機構/會員編號
Public Const ColShOrgSerialCode As Integer = 14      ' 金資序號
Public Const ColShOrgMachNumber As Integer = 15      ' 機台號碼，20240403 配合3RC多出的Column
Public Const ColShOrgTicketCode As Integer = 16      ' 票號
Public Const ColShOrgNote       As Integer = 17      ' 備註
Public Const ColShOrgChannel    As Integer = 18      ' 註記

Public Const RowShMainTSCCandRange As String = "A2:R2" ' 交易內容備用位置，3RC的位置跟解析PDF出來的位置不同

' 假設原始資料最多的 Column 數目就是最多 60 欄
Public Const MaxSrcCol         As Long = 60

' 這兩個 交易分行代碼 基本上根據 BU 的說法
' 是指 "自動化交易" 這類的分行代碼都會是0880 對 BU 來說非必要的資訊
' 所以特別獨立出來，方便後面統一過濾
Public Const SelfServiceID   As String = "0880"      ' 您幫幫您
Public Const SelfServiceID_2 As String = "880"       ' 您幫幫您

Public Const BtnStatusCleanData   As String = "清除前次分析資料"
Public Const BtnStatusStartRun    As String = "分析交易明細資料中" ' "開始分析"
Public Const BtnStatusStopRun     As String = "" ' "完成分析"
Public Const BtnStatusApplySimple As String = "準備交易明細資料"
Public Const BtnStatusStartRebase As String = "開始整理"
Public Const BtnStatusStopRebase  As String = "整理完畢"

' =========================================================
' Sheet 2.1 主頁面 所使用的常數
Public Const RowShMainContentRange As String = "F3" '"F2:F30"

' =========================================================
' Sheet 2.2 清整後資料 所使用的常數
Public Const ColShInDataTSDate     As Integer = 1        ' 交易日期
Public Const ColShInDataAccDate    As Integer = 2        ' 帳務日期
Public Const ColShInDataTSCode     As Integer = 3        ' 交易代碼
Public Const ColShInDataTSTime     As Integer = 4        ' 交易時間
Public Const ColShInDataBranchID   As Integer = 5        ' 交易分行
Public Const ColShInDataTSTeller   As Integer = 6        ' 交易櫃員
Public Const ColShInDataSummary    As Integer = 7        ' 摘要 '
Public Const ColShInDataAmount     As Integer = 8        ' 金額
Public Const ColShInDataBalance    As Integer = 9        ' 餘額
Public Const ColShInDataAccount    As Integer = 10       ' 轉出入帳號
Public Const ColShInDataMemberID   As Integer = 11       ' 合作機構/會員編號
Public Const ColShInDataSerialCode As Integer = 12       ' 金資序號
Public Const ColShInDataChannel    As Integer = 13       ' 註記
Public Const ColShInDataNote       As Integer = 14       ' 備註
Public Const ColShInDataTSMonth    As Integer = 15       ' 交易月份
Public Const ColShInDataTSSummary  As Integer = 16       ' 交易摘要
Public Const ColShInDataAmtDraw    As Integer = 17       ' 轉出金額
Public Const ColShInDataAmtDeposit As Integer = 18       ' 存入金額
Public Const ColShInDataBankCode   As Integer = 19       ' 銀行代碼
Public Const ColShInDataTSType     As Integer = 20       ' TranType
Public Const ColShInDataATMLoc     As Integer = 21       ' ATM地點
Public Const ColShInDataATMCity    As Integer = 22       ' ATM縣市
Public Const ColShInDataATMArea    As Integer = 23       ' ATM區域
Public Const ColShInDataBranchName As Integer = 24       ' 分行名
Public Const ColShInDataBranchCity As Integer = 25       ' 分行縣市
Public Const ColShInDataBranchArea As Integer = 26       ' 分行區域
Public Const ColShInDataTSLoc      As Integer = 27       ' 交易地點
Public Const ColShInDataTSChannel  As Integer = 28       ' 交易通路
Public Const ColShInDataTSOClock   As Integer = 29       ' 簡易交易時間, AC, 29
Public Const ColShInDataVAccCName  As Integer = 30       ' 照會帳戶, 虛擬帳號對應公司名, AD, 30
Public Const ColShInDataVAccReason As Integer = 31       ' 標誌虛擬原因, AE, 31
Public Const ColShInDataWAccCName  As Integer = 32       ' 警示帳戶, AF, 32
Public Const ColShInDataPAccCName  As Integer = 33       ' 樞紐顯示帳戶, AG, 32

Public Const ColShInDataAmountName     As String = "Amt" ' H, 8, ColShInDataAmount
Public Const ColShInDataAccountName    As String = "轉出入帳號" ' J, 10, ColShInDataAccount
Public Const ColShInDataTSMonthName    As String = "交易月份" ' O, 15, ColShInDataTSMonth
Public Const ColShInDataTSSummaryName  As String = "交易摘要" ' P, 16, ColShInDataTSSummary
Public Const ColShInDataBankCodeName   As String = "銀行代碼" ' S, 19, ColShInDataBankCode
Public Const ColShInDataTSTypeName     As String = "TranType" ' T, 20, ColShInDataTSType
Public Const ColShInDataATMLocName     As String = "ATM地點" ' U, 21, ColShInDataATMLoc
Public Const ColShInDataATMCityName    As String = "ATM縣市" ' V, 22, ColShInDataATMCity
Public Const ColShInDataATMAreaName    As String = "ATM區域" ' W, 23, ColShInDataATMArea
Public Const ColShInDataBrShowName     As String = "分行名" ' X, 24, ColShInDataBranchName
Public Const ColShInDataBranchCityName As String = "分行縣市" ' Y, 25, ColShInDataBranchCity
Public Const ColShInDataBranchAreaName As String = "分行區域" ' Z, 26, ColShInDataBranchArea
Public Const ColShInDataTSLocName      As String = "交易地點" ' AA, 27, ColShInDataTSLoc
Public Const ColShInDataTSChName       As String = "交易通路" ' AB, 28, ColShInDataTSChannel
Public Const ColShInDataTSOClockName   As String = "簡化交易時間" ' AC, 29, ColShInDataTSOClock
Public Const ColShInDataVAccCShowName  As String = "照會帳戶" ' AD, 30, ColShInDataVAccCName
Public Const ColShInDataVAccReasonName As String = "照會原因" ' AE, 31, ColShInDataVAccReason
Public Const ColShInDataWAccCShowName  As String = "警示帳戶" ' AF, 32, ColShInDataWAccCName
Public Const ColShInDataPAccCShowName  As String = "樞紐顯示帳戶" ' AG, 33, ColShInDataPAccCName

Public Const RowShInDataEmpty       As String = "1:5" ' 20240403 配合3RC 修改 1:7 到 1:5
Public Const ColShInDataRangePrefix As String = "A6:AG" '"A8:AD" ' 20240403 配合3RC 修改 A8:AG 到 A6:AG

Public Const ColShInDataBeginRange      As String = "A1"
Public Const ColShInDataCustomerRange   As String = "A4"
Public Const ColShInDataCustNameRange   As String = "B4"
Public Const ColShInDataAccRange        As String = "A5"
Public Const ColShInDataAccIDRange      As String = "B5"
Public Const ColShInDataProdRange       As String = "F5"
Public Const ColShInDataProdCateRange   As String = "G5"
Public Const ColShInDataCurrRange       As String = "J5"
Public Const ColShInDataCurrTypeRange   As String = "K5"
Public Const ColShInDataPrintRange      As String = "N4"
Public Const ColShInDataPrintDateRange  As String = "O4"
Public Const ColShInDataQStartRange     As String = "N5"
Public Const ColShInDataQStartDateRange As String = "O5"
Public Const ColShInDataTellerRange     As String = "Q4"
Public Const ColShInDataTellerCodeRange As String = "R4"
Public Const ColShInDataQEndRange       As String = "Q5"
Public Const ColShInDataQEndDateRange   As String = "R5"
Public Const ColShInDataTSCateRange     As String = "I2"

Public Const ColShInDataCustomerRegex As String = "戶名"
Public Const ColShInDataAccRegex      As String = "帳號"
Public Const ColShInDataTSDateRegex   As String = "交易日期"
Public Const ColShInDataProdRegex     As String = "產品別"
Public Const ColShInDataCurrRegex     As String = "幣別"
Public Const ColShInDataPrintRegex    As String = "列印日期"
Public Const ColShInDataQStartRegex   As String = "查詢起日"
Public Const ColShInDataTellerRegex   As String = "櫃員代號"
Public Const ColShInDataQEndRegex     As String = "查詢迄日"

Public Const ColShInDataChSAPostfix  As String = "系統自動"
Public Const ColShInDataChBRPostfix  As String = "臨櫃分行"
Public Const ColShInDataChNBRPostfix As String = "未知分行"
Public Const ColShInDataChWTPostfix  As String = "電匯"
Public Const ColShInDataChATMPostfix As String = "網路或ATM"

' =========================================================
' Sheet 3.1 交易明細 所使用的常數
Public Const RowShSimpleDataBegin   As Long = 9
Public Const RowShSimpleHeaderBegin As Long = (RowShSimpleDataBegin - 1)
Public Const RowShInDataSimpleGaps  As Long = RowShSimpleDataBegin - RowDataBegin
Public Const RowShSimpleEmpty       As String = "1:7"
Public Const RowShSimpleNotEmpty    As String = "7:1048576"

Public Const ColShSimpleSmallColW As Long = 13
Public Const ColShSimpleColWidth  As Long = 15

Public Const ColShSimpleTSDate      As String = "A"             ' 交易日期
Public Const ColShSimpleTSSummary   As String = "B"             ' 交易摘要
Public Const ColShSimpleTSTime      As String = "C"             ' 交易時間
Public Const ColShSimpleTSOut       As String = "D"             ' 支出, Out
Public Const ColShSimpleTSIn        As String = "E"             ' 收入, In
Public Const ColShSimpleBalance     As String = "F"             ' 餘額
Public Const ColShSimpleTSInOutAcc  As String = "G"             ' 轉出入帳號
Public Const ColShSimpleNote        As String = "H"             ' 備註
Public Const ColShSimpleChannel     As String = "I"             ' 註記
Public Const ColShSimpleTSLoc       As String = "J"             ' 交易地點
Public Const ColShSimpleAlertReason As String = "K"             ' 顏色標註理由 ' 顯示警告原因的直欄號碼

Public Const ColShSimpleLastCol     As String = ColShSimpleAlertReason

Public Const ColShSimpleBalance1K   As Double = 1000            ' 1000元
Public Const ColShSimpleTSOut100K   As Double = 100000          ' 10萬元
Public Const ColShSimpleTSOut450K   As Double = 450000          ' 45萬元
Public Const ColShSimpleTSOut500K   As Double = 500000          ' 50萬元
Public Const ColShSimpleTSIn100K    As Double = ColShSimpleTSOut100K          ' 10萬元
Public Const ColShSimpleTSIn450K    As Double = ColShSimpleTSOut450K          ' 45萬元
Public Const ColShSimpleTSIn500K    As Double = ColShSimpleTSOut500K          ' 50萬元

Public Const ColShSimpleAlertRange  As String = ColShSimpleAlertReason & "8"
Public Const ColShSimpleAlertName   As String = "顏色標註理由"
Public Const ColShSimpleRepRange    As String = "L8"
Public Const ColShSimpleRepName     As String = "表徵"
Public Const ColShSimpleResRange    As String = "M8"
Public Const ColShSimpleResName     As String = "調查結果"

Public Const ColShSimpleTempTSCh       As String = "Y"                ' 交易通路
Public Const ColShSimpleTempCode       As String = "Z"                ' 銀行代碼
Public Const ColShSimpleTempVAccName   As String = "AA"               ' 照會戶名
Public Const ColShSimpleTempVAccReason As String = "AB"               ' 照會原因
Public Const ColShSimpleTempWAccName   As String = "AC"               ' 警示帳戶

' 各式警告原因
Public Const ReasonCloseTo50w    As String = "金額接近 50 萬"
Public Const ReasonSmallTSAmount As String = "小額轉帳"
Public Const ReasonTSFast        As String = "快速進出"
Public Const ReasonTSInMorning   As String = "凌晨交易"
Public Const ReasonTSLargeAmount As String = "大額進出"
Public Const ReasonDWLAlert      As String = "往來警示帳戶"

' 上方選項的字串
Public Const UiSpecifiedAcc     As String = "指定帳戶"
Public Const UiForAllAcc        As String = "所有人"
Public Const UiDisableSearch    As String = "不使用"
Public Const UiTimeWindow1D     As String = "每日"
Public Const UiTimeWindow3D     As String = "每3日"
Public Const UiTimeWindow5D     As String = "每5日"
Public Const UiTimeWindow1M     As String = "每月"
Public Const UiOccur1Time       As String = "至少1次"
Public Const UiOccur3Time       As String = "至少3次"
Public Const UiOccur5Time       As String = "至少5次"
Public Const UiOccur10Time      As String = "至少10次"

Public Const UiTimeUnitHour  As String = "時"
Public Const UiTimeUnitDay   As String = "日"
Public Const UiTimeUnitMonth As String = "月"

Public Const UiTimeWindowString As String = UiTimeWindow1D & "," & UiTimeWindow3D & "," & UiTimeWindow5D & "," & UiTimeWindow1M
Public Const UiOccurrenceString As String = UiOccur1Time & "," & UiOccur3Time & "," & UiOccur5Time & "," & UiOccur10Time
Public Const UiOpponentString   As String = UiForAllAcc & "," & UiSpecifiedAcc
Public Const UiPatternString    As String = UiDisableSearch

Public Const UiCondiChBranch As String = ColShInDataChBRPostfix
Public Const UiCondiChwireTS As String = ColShInDataChWTPostfix
Public Const UiCondiChATMAuto As String = "ATM自動化設備"
Public Const UiCondiChBrDevices As String = "跨行機台"
Public Const UiCondiChMobile As String = "行動網"
Public Const UiCondiChOnline As String = "網銀"
Public Const UiCondiMatched  As String = "符合條件"

' =========================================================
' Sheet 3.2 交易對手
Public Const PivotTableRowStart As Long = 3
Public Const PivotTableGap      As Long = 6
Public Const PivotTableCutShow  As Long = 13

Public Const PivotTableName01   As String = "MyPivotTable1"  ' 存入交易
Public Const PivotTableName02   As String = "MyPivotTable2"  ' 支出交易
Public Const PivotTableName03   As String = "MyPivotTable3"  ' 轉帳存入交易By總流量
Public Const PivotTableName04   As String = "MyPivotTable4"  ' 轉帳支出交易By總流量
Public Const PivotTableName05   As String = "MyPivotTable5"  ' 跨行轉帳存入交易By總流量
Public Const PivotTableName06   As String = "MyPivotTable6"  ' 跨行轉帳支出交易By總流量
Public Const PivotTableName07   As String = "MyPivotTable7"  ' ATM存入交易時間
Public Const PivotTableName08   As String = "MyPivotTable8"  ' ATM存入交易地點
Public Const PivotTableName09   As String = "MyPivotTable9"  ' ATM領出交易時間
Public Const PivotTableName10   As String = "MyPivotTable10" ' ATM領出交易地點

Public Const ColShMoneyCustomerName         As String = "姓名"
Public Const ColShMoneyCustomerIDName       As String = "ID"
Public Const ColShMoneyTSModelName          As String = "交易模式"
Public Const ColShMoneyYYYYMMName           As String = "年月"
Public Const ColShMoneyTimeName             As String = "時間"
Public Const ColShMoneyCityName             As String = "縣市"
Public Const ColShMoneyTSMonthName          As String = ColShInDataTSMonthName
Public Const ColShMoneyTSSummaryName        As String = ColShInDataTSSummaryName
Public Const ColShMoneyAccountName          AS String = ColShInDataAccountName
Public Const ColShMoneyATMCityName          As String = ColShInDataATMCityName
Public Const ColShMoneyTSOClockName         As String = ColShInDataTSOClockName
Public Const ColShMoneyDepositName          As String = "存入" ' "In"
Public Const ColShMoneyWithdrawName         As String = "支出" ' "Out"
Public Const ColShMoneyTSDepositName        As String = "存入交易"
Public Const ColShMoneyTSWithdrawName       As String = "支出交易"
Public Const ColShMoneyTSInByTrafficName    As String = "轉帳存入交易By總流量"
Public Const ColShMoneyTSOutByTrafficName   As String = "轉帳支出交易By總流量"
Public Const ColShMoneyTSBrInByTrafficName  As String = "跨行轉帳存入交易By總流量"
Public Const ColShMoneyTSBrOutByTrafficName As String = "跨行轉帳支出交易By總流量"
Public Const ColShMoneyATMDepositTimeName   As String = "ATM存入交易時間"
Public Const ColShMoneyATMWithdrawTimeName  As String = "ATM領出交易時間"
Public Const ColShMoneyATMDepositLocName    As String = "ATM存入交易地點"
Public Const ColShMoneyATMWithdrawLocName   As String = "ATM領出交易地點"
Public Const ColShMoneyPivotAccountName     AS String = ColShInDataPAccCShowName

Public Const ColShMoneyCountInName   As String = "計數 - 存入"
Public Const ColShMoneyCountOutName  As String = "計數 - 支出"
Public Const ColShMoneySumIn2Name    As String = "加總 - 存入2"
Public Const ColShMoneySumOut2Name   As String = "加總 - 支出2"
Public Const ColShMoneyRatioIn3Name  As String = "佔比 - 存入3"
Public Const ColShMoneyRatioOut3Name As String = "佔比 - 支出3"

' =========================================================
' Note: this one has to be the last as the item count is not fixed
Public Const RowTopCounterparty As Long = 60

' =========================================================
' Sheet 暫存區 所使用的常數
' 這裡的常數，是為了要來畫出 2.1 主頁面 各種圖，所要暫時紀錄的 "欄位" "值" 是放暫存區哪裡
' 例如:
' RowTask1TotalCountXAxis & RowTask1TotalCountYAxis
' 就是把 分月交易總次額 的橫軸 (日期) 縱軸 (次數) 放到暫存區的 row 4 & row 5
' 切去 暫存區 就可以理解

' 任務1, 分月交易總次數
Public Const RowTask1TotalCountXAxis  As Integer = 4            ' 紀載在第 4 列
Public Const RowTask1TotalCountYAxis  As Integer = 5            ' 紀載在第 5 列
Public Const RowTask1TotalCountTitle  As String = "分月交易總次數"
Public Const RowTask1TotalCountYLabel As String = "每月次數"
Public Const RowTask1TotalCountRange  As String = "H3:O5"
' XLabel is Date

' 任務2, 分月交易總金額
' 任務3, 可疑大額提領
Public Const RowTask23TotalAmountXAxis  As Integer = 7          ' 紀載在第 7 列
Public Const RowTask23TotalAmountYAxis  As Integer = 8          ' 紀載在第 8 列
Public Const RowTask23TotalAmountTitle  As String = "分月交易總金額"
Public Const RowTask23TotalAmountYLabel As String = "每月金額"
Public Const RowTask23TotalAmountRange  As String = "Q3:X5"
' XLabel is Date

' 任務4, 當日現金提領_臨櫃 (臨櫃提款)
Public Const RowTask4WithdrawOverCounterXAxis As Integer = 10   ' 紀載在第 10 列
Public Const RowTask4WithdrawOverCounterYAxis As Integer = 11   ' 紀載在第 11 列
Public Const RowTask4WithdrawAmountLBSusAmt   As Double = ColShSimpleTSOut450K
Public Const RowTask4WithdrawAmountUBSusAmt   As Double = ColShSimpleTSOut500K

' 任務5, 當日現金提領_臨櫃 (提款總額)
Public Const RowTask5WithdrawAmountXAxis  As Integer = 13       ' 紀載在第 13 列
Public Const RowTask5WithdrawAmountYAxis  As Integer = 14       ' 紀載在第 14 列
Public Const RowTask5WithdrawAmountTitle  As String = "可疑大額提領"
Public Const RowTask5WithdrawAmountYLabel As String = "發生天數"
Public Const RowTask5WithdrawAmountRange  As String = "H7:O8"

' 任務5, 當日現金提領_臨櫃 (提款分行)
Public Const RowTask5WithdrawBranchXAxis As Integer = 16        ' 紀載在第 16 列
Public Const RowTask5WithdrawBranchYAxis As Integer = 17        ' 紀載在第 17 列
Public Const RowTask5WithdrawBranchLBCnt As Double = 2

' 任務5, 去過的分行統計
Public Const RowTask5VisitedBranchXAxis  As Integer = 19         ' 紀載在第 19 列
Public Const RowTask5VisitedBranchYAxis  As Integer = 20         ' 紀載在第 20 列
Public Const RowTask5VisitedBranchTitle  As String = "去過的分行統計"
Public Const RowTask5VisitedBranchYLabel As String = "分行交易總數"
Public Const RowTask5VisitedBranchRange  As String = "Q7:X8"
Public Const RowTask5VBCodePrefix        As String = "分行代碼: "

' 任務6, 當日現金提領/存入_ATM (快進快出)
Public Const RowTask6QuickXAxis  As Integer = 42                ' 紀載在第 42 列
Public Const RowTask6QuickYAxis  As Integer = 43                ' 紀載在第 43 列
Public Const RowTask6QuickTitle  As String = "快進快出"
Public Const RowTask6QuickYLabel As String = "統計次數"
Public Const RowTask6QuickRange  As String = "H10:X10"
Public Const RowTask6QuickLBMins As Double = 10                 ' 快進快出最短時間閥值, 分鐘

' 任務7, 當日現金提領/存入_ATM (出入帳金額比)
Public Const RowTask7DiffXAxis       As Integer = 36            ' 紀載在第 36 列
Public Const RowTask7DiffTransfer    As Integer = 37            ' 紀載在第 37 列 ' 轉帳金額Y軸的值
Public Const RowTask7DiffDeposit     As Integer = 38            ' 紀載在第 38 列 ' 存現金額Y軸的值
Public Const RowTask7DiffRatio       As Integer = 39            ' 紀載在第 39 列 ' 出入帳金額比例Y軸的值
Public Const RowTask7DiffLBRatio     As Double = 0
Public Const RowTask7DiffUBRatio     As Double = 0.03
Public Const RowTask7DiffRatioTitle  As String = "出入帳金額比例"
Public Const RowTask7DiffRatioYLabel As String = "差距比例"
Public Const RowTask7DiffRatioRange  As String = "H12:X13"

' 任務9, 當日現金提領/存入_ATM
Public Const RowTask9ATMDepositXAxis     As Integer = 22        ' 紀載在第 22 列
Public Const RowTask9ATMDepositYAxis     As Integer = 23        ' 紀載在第 23 列
Public Const RowTask9ATMDepositLoc       As Integer = 24        ' 紀載在第 24 列 ' 存現地點Y軸的值
Public Const RowTask9ATMDepositLocDetail As Integer = 25        ' 紀載在第 25 列 ' 存現地點詳細資料Y軸的值
Public Const RowTask9ATMDepositLBLoc     As Integer = 3         ' 同時段多點小額存入、一地提領，疑為車手態樣 >= 3 ATM 閥值
Public Const RowTask9ATMDepositLBCnt     As Integer = 10        ' 一天內多點存現但不具地緣性 >= 10筆 閥值
Public Const RowTask9ATMDepositTitle     As String = "多個 ATM 存現"
Public Const RowTask9ATMDepositYLabel    As String = "統計次數"
Public Const RowTask9ATMDepositRange     As String = "H15:X16"

' 任務11, 當日存款餘額
Public Const RowTask11BalanceXAxis  As Integer = 27             ' 紀載在第 27 列
Public Const RowTask11BalanceYAxis  As Integer = 28             ' 紀載在第 28 列
Public Const RowTask11BalanceTitle  As String = "日期"
Public Const RowTask11BalanceYLabel As String = "餘額"
Public Const RowTask11BalanceRange  As String = "H17:X17"
Public Const RowTask11BalanceUBCnt  As Integer = ColShSimpleBalance1K

' 任務12, 當日轉帳金額
Public Const RowTask12SmallTransferXAxis   As Integer = 30      ' 紀載在第 30 列
Public Const RowTask12SmallTransfer1Dollar As Integer = 31      ' 紀載在第 31 列 ' 一元交易Y軸的值
Public Const RowTask12SmallTransferHundred As Integer = 32      ' 紀載在第 32 列 ' 小於百元交易Y軸的值
Public Const RowTask12SmallTransferMany    As Integer = 33      ' 紀載在第 33 列 ' 跨行交易Y軸的值
Public Const RowTask12TS1DollarTitle       As String = "一元交易"
Public Const RowTask12TS1DollarYLabel      As String = "交易筆數"
Public Const RowTask12TS1DollarRange       As String = "H18:O19"
Public Const RowTask12TSHundredTitle       As String = "小於百元交易"
Public Const RowTask12TSHundredYLabel      As String = "交易筆數"
Public Const RowTask12TSHundredRange       As String = "Q18:X19"
Public Const RowTask12TSManyTitle          As String = "跨行交易"
Public Const RowTask12TSManyYLabel         As String = "交易筆數"
Public Const RowTask12TSManyRange          As String = "H20:X20"
Public Const RowTask12SmallTSLBMany        As Integer = 5       ' 跨行交易分行數閥值

Public Const RowTaskAlertPrefixByMonth   As String = "在分月交易數據中，發現在 "
Public Const RowTaskAlertPostfixByCount  As String = " 有異常的月總筆數增幅"
Public Const RowTaskAlertPostfixByAmount As String = " 有異常的月總金額增幅"

Public Const RowTaskAlertPrefixByViewer      As String = "交易數據在觀察期間中，有 "
Public Const RowTaskAlertPostfixBy1Dollar    As String = " 日發現異常的一元轉帳"
Public Const RowTaskAlertPostfixByHundred    As String = " 日發現異常的小於百元轉帳"
Public Const RowTaskAlertPostfixByMany       As String = " 日發現異常的跨行交易筆數增幅"
Public Const RowTaskAlertPostfixByDiffRatio  As String = " 次發現總出入帳金額差距甚小"
Public Const RowTaskAlertPostfixByRatio      As String = " 的總出入帳金額過小差距需要注意"
Public Const RowTaskAlertPostfixByATMDeposit As String = " 次以上 ATM 多點小額存現案例"
Public Const RowTaskAlertPostfixByATMManyLoc As String = " 個以上 ATM 不同地區存現案例"
Public Const RowTaskAlertPostfixByQuickInOut As String = " 次快進快出案例"
Public Const RowTaskAlertPostfixByAmtLess1K  As String = " 天每日存款餘額少於1000元"

Public Const RowTaskAlertPrefixByWDrawer  As String = "臨櫃提領現金交易數據在觀察期間中，以下數據需要注意: "
Public Const RowTaskAlertPostfixByWDrawer As String = ""

Public Const RowTaskGeneralPrefix    As String = "發生 "
Public Const RowTaskGeneralPostfix   As String = " 筆"
Public Const RowTaskGeneralSingleDay As String = " 單日 "
Public Const RowTaskGeneralDay As String = " 天 "
' =========================================================

Public Const ColTSSummaryVal01   As String = "01現金交易"
Public Const ColTSSummaryVal01KW As String = "現"
Public Const ColTSSummaryVal02   As String = "02匯款交易"
Public Const ColTSSummaryVal02KW As String = "匯"
Public Const ColTSSummaryVal03   As String = "03跨行轉帳"
Public Const ColTSSummaryVal03KW As String = "跨"
Public Const ColTSSummaryVal04   As String = "04本行轉帳"
Public Const ColTSSummaryVal04KW As String = "轉"
Public Const ColTSSummaryValOt   As String = "其他交易"

Public Const ColATMValOt As String = "他行ATM"

' item 19, 2024/3/27, 保留手續費與利息後，以下暫時不會用到。
'Public Const ColSummaryHandleFee As String = "手續費"
'Public Const ColSummaryInterest  As String = "利息"

Public Const ATMChannelString  As String = "存款機,ＡＴＭ,ATM"
Public Const XMLChannelString  As String = "XML,ＸＭＬ,XML"
Public Const CityNameString    As String = "基隆市, 新北市, 台北市, 桃園市, 新竹縣, 新竹市, 苗栗縣, 台中市, " & _
                                           "彰化縣, 南投縣, 雲林縣, 嘉義縣, 嘉義市, 台南市, 高雄市, 屏東縣, " & _
                                           "宜蘭縣, 花蓮縣, 台東縣, 澎湖縣, 連江縣"
Public Const EarlyMorningBegin As String = "0:00:00"
Public Const EarlyMorningEnd   As String = "4:00:00"

Public Const HintDoubleClick As String = "雙擊左邊儲存格可改排序規則"

Public Const GeneralDelimiter As String = "XX9527XX"

Public Const ColNoteChValMobile   As String = UiCondiChMobile
Public Const ColNoteChValOnline   As String = UiCondiChOnline
Public Const ColNoteChValPayment  As String = "收付網"
Public Const ColNoteChValSecurity As String = "證券款"
Public Const ColNoteChValFax      As String = "傳真銀"
Public Const ColNoteChValFEDI     As String = "FEDI"
Public Const ColNoteChValTAX      As String = "繳稅"
Public Const ColNoteChValIPASS    As String = "一卡通"
Public Const ColNoteChValCrossBR  As String = UiCondiChBrDevices

Public Const ColClerkChVal01 As String = "99998"
Public Const ColClerkChVal02 As String = "99997"

