Attribute VB_Name = "Consts"
Option Explicit

' If you cannot see the following Chinese words,
' please change your encoding from UTF-8 to Big5-HKSCS.
' �p�G�A�ݤ���o�q��r�A�бN�A�� encoding �q UTF-8 �אּ Big5-HKSCS�C

' �S�O���w�q�ȡA�Ψ� initial �@�Ӫ�array �ɵ��w����l�ȡC
Public Const EmptyArrayValue As Long = -1

' Excel style
Public Const FontSize          As Integer = 12
Public Const FontName          As String = "�L�n������"
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
Public Const FileBadAcc      As String = "ĵ�ܤ�.xlsx"
Public Const SheetNameBadAcc As String = "ĵ�ܤ�"

' Virtual account
Public Const FileVirtualAcc      As String = "�����b��.xlsx"
Public Const SheetNameVirtualAcc As String = "�ӷ|���"

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

' �D�n�� Sheet �W�r
Public Const SheetNameOrginal      As String = "1��l���"          ' �ثe�]�{���[�J
Public Const SheetNameMain         As String = "2.1�D����"          ' �`�n�X�{
Public Const SheetNameInputData    As String = "2.2�M�����"      ' �`�n�X�{
Public Const SheetNameSimple       As String = "3.1�������"        ' �`�n�X�{
Public Const SheetNameMoney        As String = "3.2���y�P������"   ' �`�n�X�{
Public Const SheetNameBranch       As String = "����M��"           ' ����
Public Const SheetNameATM          As String = "ATM�M��"           ' ����
Public Const SheetNameIntermediate As String = "�Ȧs��"             ' ����
Public Const SheetNameLabel        As String = "�ۭq�Хܳ]�w"        ' ����

' ��l��Ʊq�ĴX��}�l�~�O�u�����
Public Const RowDataBegin As Long = 7 ' 20240403, �t�X3RC�q 9 �אּ 7
Public Const RowHeaderBegin As Long = (RowDataBegin - 1)

' =========================================================
' Sheet "1��l���" �̪�����m
Public Const ColShOrgTSDate     As Integer = 1       ' ������
Public Const ColShOrgAccDate    As Integer = 2       ' �b�Ȥ��
Public Const ColShOrgTSCode     As Integer = 3       ' ����N�X
Public Const ColShOrgTSTime     As Integer = 4       ' ����ɶ�
Public Const ColShOrgBranchID   As Integer = 5       ' �������
Public Const ColShOrgTSTeller   As Integer = 6       ' ����d��
Public Const ColShOrgSummary    As Integer = 7       ' �K�n
Public Const ColShOrgAmtDraw    As Integer = 8       ' Out ' 20240403 �t�X3RC�אּ ��X
Public Const ColShOrgAmtDeposit As Integer = 9       ' In  ' 20240403 �t�X3RC�אּ �s�J
Public Const ColShOrgBalance    As Integer = 10      ' �l�B
Public Const ColShOrgBankCat    As Integer = 11      ' ��w�O�A20240403 �t�X3RC�h�X��Column
Public Const ColShOrgAccount    As Integer = 12      ' ��X�J�b��
Public Const ColShOrgMemberID   As Integer = 13      ' �X�@���c/�|���s��
Public Const ColShOrgSerialCode As Integer = 14      ' ����Ǹ�
Public Const ColShOrgMachNumber As Integer = 15      ' ���x���X�A20240403 �t�X3RC�h�X��Column
Public Const ColShOrgTicketCode As Integer = 16      ' ����
Public Const ColShOrgNote       As Integer = 17      ' �Ƶ�
Public Const ColShOrgChannel    As Integer = 18      ' ���O

Public Const RowShMainTSCCandRange As String = "A2:R2" ' ������e�ƥΦ�m�A3RC����m��ѪRPDF�X�Ӫ���m���P

' ���]��l��Ƴ̦h�� Column �ƥشN�O�̦h 60 ��
Public Const MaxSrcCol         As Long = 60

' �o��� �������N�X �򥻤W�ھ� BU �����k
' �O�� "�۰ʤƥ��" �o��������N�X���|�O0880 �� BU �ӻ��D���n����T
' �ҥH�S�O�W�ߥX�ӡA��K�᭱�Τ@�L�o
Public Const SelfServiceID   As String = "0880"      ' �z�����z
Public Const SelfServiceID_2 As String = "880"       ' �z�����z

Public Const BtnStatusCleanData   As String = "�M���e�����R���"
Public Const BtnStatusStartRun    As String = "���R������Ӹ�Ƥ�" ' "�}�l���R"
Public Const BtnStatusStopRun     As String = "" ' "�������R"
Public Const BtnStatusApplySimple As String = "�ǳƥ�����Ӹ��"
Public Const BtnStatusStartRebase As String = "�}�l��z"
Public Const BtnStatusStopRebase  As String = "��z����"

' =========================================================
' Sheet 2.1 �D���� �ҨϥΪ��`��
Public Const RowShMainContentRange As String = "F3" '"F2:F30"

' =========================================================
' Sheet 2.2 �M����� �ҨϥΪ��`��
Public Const ColShInDataTSDate     As Integer = 1        ' ������
Public Const ColShInDataAccDate    As Integer = 2        ' �b�Ȥ��
Public Const ColShInDataTSCode     As Integer = 3        ' ����N�X
Public Const ColShInDataTSTime     As Integer = 4        ' ����ɶ�
Public Const ColShInDataBranchID   As Integer = 5        ' �������
Public Const ColShInDataTSTeller   As Integer = 6        ' ����d��
Public Const ColShInDataSummary    As Integer = 7        ' �K�n '
Public Const ColShInDataAmount     As Integer = 8        ' ���B
Public Const ColShInDataBalance    As Integer = 9        ' �l�B
Public Const ColShInDataAccount    As Integer = 10       ' ��X�J�b��
Public Const ColShInDataMemberID   As Integer = 11       ' �X�@���c/�|���s��
Public Const ColShInDataSerialCode As Integer = 12       ' ����Ǹ�
Public Const ColShInDataChannel    As Integer = 13       ' ���O
Public Const ColShInDataNote       As Integer = 14       ' �Ƶ�
Public Const ColShInDataTSMonth    As Integer = 15       ' ������
Public Const ColShInDataTSSummary  As Integer = 16       ' ����K�n
Public Const ColShInDataAmtDraw    As Integer = 17       ' ��X���B
Public Const ColShInDataAmtDeposit As Integer = 18       ' �s�J���B
Public Const ColShInDataBankCode   As Integer = 19       ' �Ȧ�N�X
Public Const ColShInDataTSType     As Integer = 20       ' TranType
Public Const ColShInDataATMLoc     As Integer = 21       ' ATM�a�I
Public Const ColShInDataATMCity    As Integer = 22       ' ATM����
Public Const ColShInDataATMArea    As Integer = 23       ' ATM�ϰ�
Public Const ColShInDataBranchName As Integer = 24       ' ����W
Public Const ColShInDataBranchCity As Integer = 25       ' ���濤��
Public Const ColShInDataBranchArea As Integer = 26       ' ����ϰ�
Public Const ColShInDataTSLoc      As Integer = 27       ' ����a�I
Public Const ColShInDataTSChannel  As Integer = 28       ' ����q��
Public Const ColShInDataTSOClock   As Integer = 29       ' ²������ɶ�, AC, 29
Public Const ColShInDataVAccCName  As Integer = 30       ' �ӷ|�b��, �����b���������q�W, AD, 30
Public Const ColShInDataVAccReason As Integer = 31       ' �лx������], AE, 31
Public Const ColShInDataWAccCName  As Integer = 32       ' ĵ�ܱb��, AF, 32
Public Const ColShInDataPAccCName  As Integer = 33       ' �ϯ���ܱb��, AG, 32

Public Const ColShInDataAmountName     As String = "Amt" ' H, 8, ColShInDataAmount
Public Const ColShInDataAccountName    As String = "��X�J�b��" ' J, 10, ColShInDataAccount
Public Const ColShInDataTSMonthName    As String = "������" ' O, 15, ColShInDataTSMonth
Public Const ColShInDataTSSummaryName  As String = "����K�n" ' P, 16, ColShInDataTSSummary
Public Const ColShInDataBankCodeName   As String = "�Ȧ�N�X" ' S, 19, ColShInDataBankCode
Public Const ColShInDataTSTypeName     As String = "TranType" ' T, 20, ColShInDataTSType
Public Const ColShInDataATMLocName     As String = "ATM�a�I" ' U, 21, ColShInDataATMLoc
Public Const ColShInDataATMCityName    As String = "ATM����" ' V, 22, ColShInDataATMCity
Public Const ColShInDataATMAreaName    As String = "ATM�ϰ�" ' W, 23, ColShInDataATMArea
Public Const ColShInDataBrShowName     As String = "����W" ' X, 24, ColShInDataBranchName
Public Const ColShInDataBranchCityName As String = "���濤��" ' Y, 25, ColShInDataBranchCity
Public Const ColShInDataBranchAreaName As String = "����ϰ�" ' Z, 26, ColShInDataBranchArea
Public Const ColShInDataTSLocName      As String = "����a�I" ' AA, 27, ColShInDataTSLoc
Public Const ColShInDataTSChName       As String = "����q��" ' AB, 28, ColShInDataTSChannel
Public Const ColShInDataTSOClockName   As String = "²�ƥ���ɶ�" ' AC, 29, ColShInDataTSOClock
Public Const ColShInDataVAccCShowName  As String = "�ӷ|�b��" ' AD, 30, ColShInDataVAccCName
Public Const ColShInDataVAccReasonName As String = "�ӷ|��]" ' AE, 31, ColShInDataVAccReason
Public Const ColShInDataWAccCShowName  As String = "ĵ�ܱb��" ' AF, 32, ColShInDataWAccCName
Public Const ColShInDataPAccCShowName  As String = "�ϯ���ܱb��" ' AG, 33, ColShInDataPAccCName

Public Const RowShInDataEmpty       As String = "1:5" ' 20240403 �t�X3RC �ק� 1:7 �� 1:5
Public Const ColShInDataRangePrefix As String = "A6:AG" '"A8:AD" ' 20240403 �t�X3RC �ק� A8:AG �� A6:AG

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

Public Const ColShInDataCustomerRegex As String = "��W"
Public Const ColShInDataAccRegex      As String = "�b��"
Public Const ColShInDataTSDateRegex   As String = "������"
Public Const ColShInDataProdRegex     As String = "���~�O"
Public Const ColShInDataCurrRegex     As String = "���O"
Public Const ColShInDataPrintRegex    As String = "�C�L���"
Public Const ColShInDataQStartRegex   As String = "�d�߰_��"
Public Const ColShInDataTellerRegex   As String = "�d���N��"
Public Const ColShInDataQEndRegex     As String = "�d�ߨ���"

Public Const ColShInDataChSAPostfix  As String = "�t�Φ۰�"
Public Const ColShInDataChBRPostfix  As String = "�{�d����"
Public Const ColShInDataChNBRPostfix As String = "��������"
Public Const ColShInDataChWTPostfix  As String = "�q��"
Public Const ColShInDataChATMPostfix As String = "������ATM"

' =========================================================
' Sheet 3.1 ������� �ҨϥΪ��`��
Public Const RowShSimpleDataBegin   As Long = 9
Public Const RowShSimpleHeaderBegin As Long = (RowShSimpleDataBegin - 1)
Public Const RowShInDataSimpleGaps  As Long = RowShSimpleDataBegin - RowDataBegin
Public Const RowShSimpleEmpty       As String = "1:7"
Public Const RowShSimpleNotEmpty    As String = "7:1048576"

Public Const ColShSimpleSmallColW As Long = 13
Public Const ColShSimpleColWidth  As Long = 15

Public Const ColShSimpleTSDate      As String = "A"             ' ������
Public Const ColShSimpleTSSummary   As String = "B"             ' ����K�n
Public Const ColShSimpleTSTime      As String = "C"             ' ����ɶ�
Public Const ColShSimpleTSOut       As String = "D"             ' ��X, Out
Public Const ColShSimpleTSIn        As String = "E"             ' ���J, In
Public Const ColShSimpleBalance     As String = "F"             ' �l�B
Public Const ColShSimpleTSInOutAcc  As String = "G"             ' ��X�J�b��
Public Const ColShSimpleNote        As String = "H"             ' �Ƶ�
Public Const ColShSimpleChannel     As String = "I"             ' ���O
Public Const ColShSimpleTSLoc       As String = "J"             ' ����a�I
Public Const ColShSimpleAlertReason As String = "K"             ' �C��е��z�� ' ���ĵ�i��]�����渹�X

Public Const ColShSimpleLastCol     As String = ColShSimpleAlertReason

Public Const ColShSimpleBalance1K   As Double = 1000            ' 1000��
Public Const ColShSimpleTSOut100K   As Double = 100000          ' 10�U��
Public Const ColShSimpleTSOut450K   As Double = 450000          ' 45�U��
Public Const ColShSimpleTSOut500K   As Double = 500000          ' 50�U��
Public Const ColShSimpleTSIn100K    As Double = ColShSimpleTSOut100K          ' 10�U��
Public Const ColShSimpleTSIn450K    As Double = ColShSimpleTSOut450K          ' 45�U��
Public Const ColShSimpleTSIn500K    As Double = ColShSimpleTSOut500K          ' 50�U��

Public Const ColShSimpleAlertRange  As String = ColShSimpleAlertReason & "8"
Public Const ColShSimpleAlertName   As String = "�C��е��z��"
Public Const ColShSimpleRepRange    As String = "L8"
Public Const ColShSimpleRepName     As String = "��x"
Public Const ColShSimpleResRange    As String = "M8"
Public Const ColShSimpleResName     As String = "�լd���G"

Public Const ColShSimpleTempTSCh       As String = "Y"                ' ����q��
Public Const ColShSimpleTempCode       As String = "Z"                ' �Ȧ�N�X
Public Const ColShSimpleTempVAccName   As String = "AA"               ' �ӷ|��W
Public Const ColShSimpleTempVAccReason As String = "AB"               ' �ӷ|��]
Public Const ColShSimpleTempWAccName   As String = "AC"               ' ĵ�ܱb��

' �U��ĵ�i��]
Public Const ReasonCloseTo50w    As String = "���B���� 50 �U"
Public Const ReasonSmallTSAmount As String = "�p�B��b"
Public Const ReasonTSFast        As String = "�ֳt�i�X"
Public Const ReasonTSInMorning   As String = "�����"
Public Const ReasonTSLargeAmount As String = "�j�B�i�X"
Public Const ReasonDWLAlert      As String = "����ĵ�ܱb��"

' �W��ﶵ���r��
Public Const UiSpecifiedAcc     As String = "���w�b��"
Public Const UiForAllAcc        As String = "�Ҧ��H"
Public Const UiDisableSearch    As String = "���ϥ�"
Public Const UiTimeWindow1D     As String = "�C��"
Public Const UiTimeWindow3D     As String = "�C3��"
Public Const UiTimeWindow5D     As String = "�C5��"
Public Const UiTimeWindow1M     As String = "�C��"
Public Const UiOccur1Time       As String = "�ܤ�1��"
Public Const UiOccur3Time       As String = "�ܤ�3��"
Public Const UiOccur5Time       As String = "�ܤ�5��"
Public Const UiOccur10Time      As String = "�ܤ�10��"

Public Const UiTimeUnitHour  As String = "��"
Public Const UiTimeUnitDay   As String = "��"
Public Const UiTimeUnitMonth As String = "��"

Public Const UiTimeWindowString As String = UiTimeWindow1D & "," & UiTimeWindow3D & "," & UiTimeWindow5D & "," & UiTimeWindow1M
Public Const UiOccurrenceString As String = UiOccur1Time & "," & UiOccur3Time & "," & UiOccur5Time & "," & UiOccur10Time
Public Const UiOpponentString   As String = UiForAllAcc & "," & UiSpecifiedAcc
Public Const UiPatternString    As String = UiDisableSearch

Public Const UiCondiChBranch As String = ColShInDataChBRPostfix
Public Const UiCondiChwireTS As String = ColShInDataChWTPostfix
Public Const UiCondiChATMAuto As String = "ATM�۰ʤƳ]��"
Public Const UiCondiChBrDevices As String = "�����x"
Public Const UiCondiChMobile As String = "��ʺ�"
Public Const UiCondiChOnline As String = "����"
Public Const UiCondiMatched  As String = "�ŦX����"

' =========================================================
' Sheet 3.2 ������
Public Const PivotTableRowStart As Long = 3
Public Const PivotTableGap      As Long = 6
Public Const PivotTableCutShow  As Long = 13

Public Const PivotTableName01   As String = "MyPivotTable1"  ' �s�J���
Public Const PivotTableName02   As String = "MyPivotTable2"  ' ��X���
Public Const PivotTableName03   As String = "MyPivotTable3"  ' ��b�s�J���By�`�y�q
Public Const PivotTableName04   As String = "MyPivotTable4"  ' ��b��X���By�`�y�q
Public Const PivotTableName05   As String = "MyPivotTable5"  ' �����b�s�J���By�`�y�q
Public Const PivotTableName06   As String = "MyPivotTable6"  ' �����b��X���By�`�y�q
Public Const PivotTableName07   As String = "MyPivotTable7"  ' ATM�s�J����ɶ�
Public Const PivotTableName08   As String = "MyPivotTable8"  ' ATM�s�J����a�I
Public Const PivotTableName09   As String = "MyPivotTable9"  ' ATM��X����ɶ�
Public Const PivotTableName10   As String = "MyPivotTable10" ' ATM��X����a�I

Public Const ColShMoneyCustomerName         As String = "�m�W"
Public Const ColShMoneyCustomerIDName       As String = "ID"
Public Const ColShMoneyTSModelName          As String = "����Ҧ�"
Public Const ColShMoneyYYYYMMName           As String = "�~��"
Public Const ColShMoneyTimeName             As String = "�ɶ�"
Public Const ColShMoneyCityName             As String = "����"
Public Const ColShMoneyTSMonthName          As String = ColShInDataTSMonthName
Public Const ColShMoneyTSSummaryName        As String = ColShInDataTSSummaryName
Public Const ColShMoneyAccountName          AS String = ColShInDataAccountName
Public Const ColShMoneyATMCityName          As String = ColShInDataATMCityName
Public Const ColShMoneyTSOClockName         As String = ColShInDataTSOClockName
Public Const ColShMoneyDepositName          As String = "�s�J" ' "In"
Public Const ColShMoneyWithdrawName         As String = "��X" ' "Out"
Public Const ColShMoneyTSDepositName        As String = "�s�J���"
Public Const ColShMoneyTSWithdrawName       As String = "��X���"
Public Const ColShMoneyTSInByTrafficName    As String = "��b�s�J���By�`�y�q"
Public Const ColShMoneyTSOutByTrafficName   As String = "��b��X���By�`�y�q"
Public Const ColShMoneyTSBrInByTrafficName  As String = "�����b�s�J���By�`�y�q"
Public Const ColShMoneyTSBrOutByTrafficName As String = "�����b��X���By�`�y�q"
Public Const ColShMoneyATMDepositTimeName   As String = "ATM�s�J����ɶ�"
Public Const ColShMoneyATMWithdrawTimeName  As String = "ATM��X����ɶ�"
Public Const ColShMoneyATMDepositLocName    As String = "ATM�s�J����a�I"
Public Const ColShMoneyATMWithdrawLocName   As String = "ATM��X����a�I"
Public Const ColShMoneyPivotAccountName     AS String = ColShInDataPAccCShowName

Public Const ColShMoneyCountInName   As String = "�p�� - �s�J"
Public Const ColShMoneyCountOutName  As String = "�p�� - ��X"
Public Const ColShMoneySumIn2Name    As String = "�[�` - �s�J2"
Public Const ColShMoneySumOut2Name   As String = "�[�` - ��X2"
Public Const ColShMoneyRatioIn3Name  As String = "���� - �s�J3"
Public Const ColShMoneyRatioOut3Name As String = "���� - ��X3"

' =========================================================
' Note: this one has to be the last as the item count is not fixed
Public Const RowTopCounterparty As Long = 60

' =========================================================
' Sheet �Ȧs�� �ҨϥΪ��`��
' �o�̪��`�ơA�O���F�n�ӵe�X 2.1 �D���� �U�عϡA�ҭn�Ȯɬ����� "���" "��" �O��Ȧs�ϭ���
' �Ҧp:
' RowTask1TotalCountXAxis & RowTask1TotalCountYAxis
' �N�O�� �������`���B ����b (���) �a�b (����) ���Ȧs�Ϫ� row 4 & row 5
' ���h �Ȧs�� �N�i�H�z��

' ����1, �������`����
Public Const RowTask1TotalCountXAxis  As Integer = 4            ' �����b�� 4 �C
Public Const RowTask1TotalCountYAxis  As Integer = 5            ' �����b�� 5 �C
Public Const RowTask1TotalCountTitle  As String = "�������`����"
Public Const RowTask1TotalCountYLabel As String = "�C�릸��"
Public Const RowTask1TotalCountRange  As String = "H3:O5"
' XLabel is Date

' ����2, �������`���B
' ����3, �i�äj�B����
Public Const RowTask23TotalAmountXAxis  As Integer = 7          ' �����b�� 7 �C
Public Const RowTask23TotalAmountYAxis  As Integer = 8          ' �����b�� 8 �C
Public Const RowTask23TotalAmountTitle  As String = "�������`���B"
Public Const RowTask23TotalAmountYLabel As String = "�C����B"
Public Const RowTask23TotalAmountRange  As String = "Q3:X5"
' XLabel is Date

' ����4, ���{������_�{�d (�{�d����)
Public Const RowTask4WithdrawOverCounterXAxis As Integer = 10   ' �����b�� 10 �C
Public Const RowTask4WithdrawOverCounterYAxis As Integer = 11   ' �����b�� 11 �C
Public Const RowTask4WithdrawAmountLBSusAmt   As Double = ColShSimpleTSOut450K
Public Const RowTask4WithdrawAmountUBSusAmt   As Double = ColShSimpleTSOut500K

' ����5, ���{������_�{�d (�����`�B)
Public Const RowTask5WithdrawAmountXAxis  As Integer = 13       ' �����b�� 13 �C
Public Const RowTask5WithdrawAmountYAxis  As Integer = 14       ' �����b�� 14 �C
Public Const RowTask5WithdrawAmountTitle  As String = "�i�äj�B����"
Public Const RowTask5WithdrawAmountYLabel As String = "�o�ͤѼ�"
Public Const RowTask5WithdrawAmountRange  As String = "H7:O8"

' ����5, ���{������_�{�d (���ڤ���)
Public Const RowTask5WithdrawBranchXAxis As Integer = 16        ' �����b�� 16 �C
Public Const RowTask5WithdrawBranchYAxis As Integer = 17        ' �����b�� 17 �C
Public Const RowTask5WithdrawBranchLBCnt As Double = 2

' ����5, �h�L������έp
Public Const RowTask5VisitedBranchXAxis  As Integer = 19         ' �����b�� 19 �C
Public Const RowTask5VisitedBranchYAxis  As Integer = 20         ' �����b�� 20 �C
Public Const RowTask5VisitedBranchTitle  As String = "�h�L������έp"
Public Const RowTask5VisitedBranchYLabel As String = "�������`��"
Public Const RowTask5VisitedBranchRange  As String = "Q7:X8"
Public Const RowTask5VBCodePrefix        As String = "����N�X: "

' ����6, ���{������/�s�J_ATM (�ֶi�֥X)
Public Const RowTask6QuickXAxis  As Integer = 42                ' �����b�� 42 �C
Public Const RowTask6QuickYAxis  As Integer = 43                ' �����b�� 43 �C
Public Const RowTask6QuickTitle  As String = "�ֶi�֥X"
Public Const RowTask6QuickYLabel As String = "�έp����"
Public Const RowTask6QuickRange  As String = "H10:X10"
Public Const RowTask6QuickLBMins As Double = 10                 ' �ֶi�֥X�̵u�ɶ��֭�, ����

' ����7, ���{������/�s�J_ATM (�X�J�b���B��)
Public Const RowTask7DiffXAxis       As Integer = 36            ' �����b�� 36 �C
Public Const RowTask7DiffTransfer    As Integer = 37            ' �����b�� 37 �C ' ��b���BY�b����
Public Const RowTask7DiffDeposit     As Integer = 38            ' �����b�� 38 �C ' �s�{���BY�b����
Public Const RowTask7DiffRatio       As Integer = 39            ' �����b�� 39 �C ' �X�J�b���B���Y�b����
Public Const RowTask7DiffLBRatio     As Double = 0
Public Const RowTask7DiffUBRatio     As Double = 0.03
Public Const RowTask7DiffRatioTitle  As String = "�X�J�b���B���"
Public Const RowTask7DiffRatioYLabel As String = "�t�Z���"
Public Const RowTask7DiffRatioRange  As String = "H12:X13"

' ����9, ���{������/�s�J_ATM
Public Const RowTask9ATMDepositXAxis     As Integer = 22        ' �����b�� 22 �C
Public Const RowTask9ATMDepositYAxis     As Integer = 23        ' �����b�� 23 �C
Public Const RowTask9ATMDepositLoc       As Integer = 24        ' �����b�� 24 �C ' �s�{�a�IY�b����
Public Const RowTask9ATMDepositLocDetail As Integer = 25        ' �����b�� 25 �C ' �s�{�a�I�ԲӸ��Y�b����
Public Const RowTask9ATMDepositLBLoc     As Integer = 3         ' �P�ɬq�h�I�p�B�s�J�B�@�a����A�ì�����A�� >= 3 ATM �֭�
Public Const RowTask9ATMDepositLBCnt     As Integer = 10        ' �@�Ѥ��h�I�s�{������a�t�� >= 10�� �֭�
Public Const RowTask9ATMDepositTitle     As String = "�h�� ATM �s�{"
Public Const RowTask9ATMDepositYLabel    As String = "�έp����"
Public Const RowTask9ATMDepositRange     As String = "H15:X16"

' ����11, ���s�ھl�B
Public Const RowTask11BalanceXAxis  As Integer = 27             ' �����b�� 27 �C
Public Const RowTask11BalanceYAxis  As Integer = 28             ' �����b�� 28 �C
Public Const RowTask11BalanceTitle  As String = "���"
Public Const RowTask11BalanceYLabel As String = "�l�B"
Public Const RowTask11BalanceRange  As String = "H17:X17"
Public Const RowTask11BalanceUBCnt  As Integer = ColShSimpleBalance1K

' ����12, �����b���B
Public Const RowTask12SmallTransferXAxis   As Integer = 30      ' �����b�� 30 �C
Public Const RowTask12SmallTransfer1Dollar As Integer = 31      ' �����b�� 31 �C ' �@�����Y�b����
Public Const RowTask12SmallTransferHundred As Integer = 32      ' �����b�� 32 �C ' �p��ʤ����Y�b����
Public Const RowTask12SmallTransferMany    As Integer = 33      ' �����b�� 33 �C ' �����Y�b����
Public Const RowTask12TS1DollarTitle       As String = "�@�����"
Public Const RowTask12TS1DollarYLabel      As String = "�������"
Public Const RowTask12TS1DollarRange       As String = "H18:O19"
Public Const RowTask12TSHundredTitle       As String = "�p��ʤ����"
Public Const RowTask12TSHundredYLabel      As String = "�������"
Public Const RowTask12TSHundredRange       As String = "Q18:X19"
Public Const RowTask12TSManyTitle          As String = "�����"
Public Const RowTask12TSManyYLabel         As String = "�������"
Public Const RowTask12TSManyRange          As String = "H20:X20"
Public Const RowTask12SmallTSLBMany        As Integer = 5       ' ���������ƻ֭�

Public Const RowTaskAlertPrefixByMonth   As String = "�b�������ƾڤ��A�o�{�b "
Public Const RowTaskAlertPostfixByCount  As String = " �����`�����`���ƼW�T"
Public Const RowTaskAlertPostfixByAmount As String = " �����`�����`���B�W�T"

Public Const RowTaskAlertPrefixByViewer      As String = "����ƾڦb�[��������A�� "
Public Const RowTaskAlertPostfixBy1Dollar    As String = " ��o�{���`���@����b"
Public Const RowTaskAlertPostfixByHundred    As String = " ��o�{���`���p��ʤ���b"
Public Const RowTaskAlertPostfixByMany       As String = " ��o�{���`����������ƼW�T"
Public Const RowTaskAlertPostfixByDiffRatio  As String = " ���o�{�`�X�J�b���B�t�Z�Ƥp"
Public Const RowTaskAlertPostfixByRatio      As String = " ���`�X�J�b���B�L�p�t�Z�ݭn�`�N"
Public Const RowTaskAlertPostfixByATMDeposit As String = " ���H�W ATM �h�I�p�B�s�{�ר�"
Public Const RowTaskAlertPostfixByATMManyLoc As String = " �ӥH�W ATM ���P�a�Ϧs�{�ר�"
Public Const RowTaskAlertPostfixByQuickInOut As String = " ���ֶi�֥X�ר�"
Public Const RowTaskAlertPostfixByAmtLess1K  As String = " �ѨC��s�ھl�B�֩�1000��"

Public Const RowTaskAlertPrefixByWDrawer  As String = "�{�d����{������ƾڦb�[��������A�H�U�ƾڻݭn�`�N: "
Public Const RowTaskAlertPostfixByWDrawer As String = ""

Public Const RowTaskGeneralPrefix    As String = "�o�� "
Public Const RowTaskGeneralPostfix   As String = " ��"
Public Const RowTaskGeneralSingleDay As String = " ��� "
Public Const RowTaskGeneralDay As String = " �� "
' =========================================================

Public Const ColTSSummaryVal01   As String = "01�{�����"
Public Const ColTSSummaryVal01KW As String = "�{"
Public Const ColTSSummaryVal02   As String = "02�״ڥ��"
Public Const ColTSSummaryVal02KW As String = "��"
Public Const ColTSSummaryVal03   As String = "03�����b"
Public Const ColTSSummaryVal03KW As String = "��"
Public Const ColTSSummaryVal04   As String = "04������b"
Public Const ColTSSummaryVal04KW As String = "��"
Public Const ColTSSummaryValOt   As String = "��L���"

Public Const ColATMValOt As String = "�L��ATM"

' item 19, 2024/3/27, �O�d����O�P�Q����A�H�U�Ȯɤ��|�Ψ�C
'Public Const ColSummaryHandleFee As String = "����O"
'Public Const ColSummaryInterest  As String = "�Q��"

Public Const ATMChannelString  As String = "�s�ھ�,�Ϣ��,ATM"
Public Const XMLChannelString  As String = "XML,��ۢ�,XML"
Public Const CityNameString    As String = "�򶩥�, �s�_��, �x�_��, ��饫, �s�˿�, �s�˥�, �]�߿�, �x����, " & _
                                           "���ƿ�, �n�뿤, ���L��, �Ÿq��, �Ÿq��, �x�n��, ������, �̪F��, " & _
                                           "�y����, �Ὤ��, �x�F��, ���, �s����"
Public Const EarlyMorningBegin As String = "0:00:00"
Public Const EarlyMorningEnd   As String = "4:00:00"

Public Const HintDoubleClick As String = "���������x�s��i��ƧǳW�h"

Public Const GeneralDelimiter As String = "XX9527XX"

Public Const ColNoteChValMobile   As String = UiCondiChMobile
Public Const ColNoteChValOnline   As String = UiCondiChOnline
Public Const ColNoteChValPayment  As String = "���I��"
Public Const ColNoteChValSecurity As String = "�Ҩ��"
Public Const ColNoteChValFax      As String = "�ǯu��"
Public Const ColNoteChValFEDI     As String = "FEDI"
Public Const ColNoteChValTAX      As String = "ú�|"
Public Const ColNoteChValIPASS    As String = "�@�d�q"
Public Const ColNoteChValCrossBR  As String = UiCondiChBrDevices

Public Const ColClerkChVal01 As String = "99998"
Public Const ColClerkChVal02 As String = "99997"

