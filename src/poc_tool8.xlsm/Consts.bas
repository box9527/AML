Attribute VB_Name = "Consts"
Option Explicit

' If you cannot see the following Chinese words,
' please change your encoding from UTF-8 to Big5-HKSCS.
' �p�G�A�ݤ���o�q��r�A�бN�A�� encoding �q UTF-8 �אּ Big5-HKSCS�C

' �S�O���w�q�ȡA�Ψ� initial �@�Ӫ�array �ɵ��w����l�ȡC
Public Const EmptyArrayValue As Long = -1

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
Public Const ColorYellow2   As Long = &H86FFFE

' �D�n�� Sheet �W�r
Public Const SheetNameOrginal      As String = "1��l���"          ' �ثe�]�{���[�J
Public Const SheetNameMain         As String = "2.1�D����"          ' �`�n�X�{
Public Const SheetNameInputData    As String = "2.2�M�����"      ' �`�n�X�{
Public Const SheetNameSimple       As String = "3.1�������"        ' �`�n�X�{
Public Const SheetNameMoney        As String = "3.2���y�P������"   ' �`�n�X�{
Public Const SheetNameBranch       As String = "����M��"           ' ����
Public Const SheetNameIntermediate As String = "�Ȧs��"             ' ����
Public Const SheetNameLabel        As String = "�ۭq�Хܳ]�w"        ' ����

' ��l��Ʊq�ĴX��}�l�~�O�u�����
Public Const RowDataBegin       As Integer = 9

' Sheet "�M�����" �̪�����m
Public Const ColTSDate         As Integer = 1        ' ������
Public Const ColAccDate        As Integer = 2        ' �b�Ȥ��
Public Const ColTSCode         As Integer = 3        ' ����N�X
Public Const ColTSTime         As Integer = 4        ' ����ɶ�
Public Const ColBranchID       As Integer = 5        ' �������
Public Const ColTSTeller       As Integer = 6        ' ����d��
Public Const ColSummary        As Integer = 7        ' �K�n
Public Const ColAmount         As Integer = 8        ' ���B
Public Const ColBalance        As Integer = 9        ' �l�B
Public Const ColAccount        As Integer = 10       ' ��X�J�b��
Public Const ColMemberID       As Integer = 11       ' �X�@���c/�|���s��
Public Const ColSerialCode     As Integer = 12       ' ����Ǹ�
Public Const ColChannel        As Integer = 13       ' ���O
Public Const ColNote           As Integer = 14       ' �Ƶ�
Public Const ColTSMonth        As Integer = 15       ' ������
Public Const ColTSSummary      As Integer = 16       ' ����K�n
Public Const ColAmountTransfer As Integer = 17       ' ��X���B
Public Const ColAmountDeposit  As Integer = 18       ' �s�J���B
Public Const ColBankCode       As Integer = 19       ' �Ȧ�N�X
Public Const ColTSType         As Integer = 20       ' TranType
Public Const ColATMLoc         As Integer = 21       ' ATM�a�I
Public Const ColATMCity        As Integer = 22       ' ATM����
Public Const ColATMArea        As Integer = 23       ' ATM�ϰ�
Public Const ColBranchName     As Integer = 24       ' ����W
Public Const ColBranchCity     As Integer = 25       ' ���濤��
Public Const ColBranchArea     As Integer = 26       ' ����ϰ�
Public Const ColTSLoc          As Integer = 27       ' ����a�I
Public Const ColTSChannel      As Integer = 28       ' ����q��
Public Const ColTSOClock       As Integer = 29       ' ²������ɶ�

'���]��l��Ƴ̦h�� Column �ƥشN�O�̦h 50 ��
Public Const MaxSrcCol         As Integer = 50

' �o��� �������N�X �򥻤W�ھ� BU �����k
' �O�� "�۰ʤƥ��" �o��������N�X���|�O0880 �� BU �ӻ��D���n����T
' �ҥH�S�O�W�ߥX�ӡA��K�᭱�Τ@�L�o
Public Const SelfServiceID   As String = "0880"      ' �z�����z
Public Const SelfServiceID_2 As String = "880"       ' �z�����z

' =========================================================
' Sheet 3.1 ������� �ҨϥΪ��`��
' ���ĵ�i��]�����渹�X
Public Const ColAlertReason As String = "K"

' �U��ĵ�i��]
Public Const ReasonCloseTo50w     As String = "���B���� 50 �U"
Public Const ReasonSmallTSAmount  As String = "�p�B��b"
Public Const ReasonTSFast         As String = "�ֳt�i�X"
Public Const ReasonTSInMorning    As String = "�����"
Public Const ReasonTSLargeAmount  As String = "�j�B�i�X"
Public Const ReasonDWLAlert       As String = "����ĵ�ܱb��"
' =========================================================

' =========================================================
' Sheet �Ȧs�� �ҨϥΪ��`��
' �o�̪��`�ơA�O���F�n�ӵe�X 2.1 �D���� �U�عϡA�ҭn�Ȯɬ����� "���" "��" �O��Ȧs�ϭ���
' �Ҧp:
' RowTask1TotalCountXAxis & RowTask1TotalCountYAxis 
' �N�O�� �������`���B ����b (���) �a�b (����) ���Ȧs�Ϫ� row 4 & row 5
' ���h �Ȧs�� �N�i�H�z��

' ����1, �������`����
Public Const RowTask1TotalCountXAxis As Integer = 4             ' �����b�� 4 �C
Public Const RowTask1TotalCountYAxis As Integer = 5             ' �����b�� 5 �C

' ����2, �������`���B
' ����3, �i�äj�B����
Public Const RowTask23TotalAmountXAxis As Integer = 7           ' �����b�� 7 �C
Public Const RowTask23TotalAmountYAxis As Integer = 8           ' �����b�� 8 �C

' ����4, ���{������_�{�d (�{�d����)
Public Const RowTask4WithdrawOverCounterXAxis As Integer = 10   ' �����b�� 10 �C
Public Const RowTask4WithdrawOverCounterYAxis As Integer = 11   ' �����b�� 11 �C

' ����5, ���{������_�{�d (�����`�B)
Public Const RowTask5WithdrawSummaryXAxis As Integer = 13       ' �����b�� 13 �C
Public Const RowTask5WithdrawSummaryYAxis As Integer = 14       ' �����b�� 14 �C

' ����5, ���{������_�{�d (���ڤ���)
Public Const RowTask5WithdrawBranchXAxis As Integer = 16        ' �����b�� 16 �C
Public Const RowTask5WithdrawBranchYAxis As Integer = 17        ' �����b�� 17 �C

' ����5, �h�L������έp
Public Const RowTask5VisitedBranchXAxis As Integer = 19         ' �����b�� 19 �C
Public Const RowTask5VisitedBranchYAxis As Integer = 20         ' �����b�� 20 �C

' ����6, ���{������/�s�J_ATM (�ֶi�֥X)
Public Const RowTask6QuickXAxis As Integer = 42                 ' �����b�� 42 �C
Public Const RowTask6QuickYAxis As Integer = 43                 ' �����b�� 43 �C

' ����7, ���{������/�s�J_ATM (�X�J�b���B��)
Public Const RowTask7DiffXAxis    As Integer = 36               ' �����b�� 36 �C
Public Const RowTask7DiffTransfer As Integer = 37               ' �����b�� 37 �C ' ��b���BY�b����
Public Const RowTask7DiffDeposit  As Integer = 38               ' �����b�� 38 �C ' �s�{���BY�b����
Public Const RowTask7DiffRatio    As Integer = 39               ' �����b�� 39 �C ' �X�J�b���B���Y�b����

' ����9, ���{������/�s�J_ATM
Public Const RowTask9ATMDepositXAxis As Integer = 22            ' �����b�� 22 �C
Public Const RowTask9ATMDepositYAxis As Integer = 23            ' �����b�� 23 �C
Public Const RowTask9ATMDepositLocation As Integer = 24         ' �����b�� 24 �C ' �s�{�a�IY�b����
Public Const RowTask9ATMDepositLocationDetail As Integer = 25   ' �����b�� 25 �C ' �s�{�a�I�ԲӸ��Y�b����

' ����11, ���s�ھl�B
Public Const RowTask11BalanceXAxis As Integer = 27              ' �����b�� 27 �C
Public Const RowTask11BalanceYAxis As Integer = 28              ' �����b�� 28 �C

' ����12, �����b���B
Public Const RowTask12SmallTransferXAxis     As Integer = 30    ' �����b�� 30 �C
Public Const RowTask12SmallTransferOneDollar As Integer = 31    ' �����b�� 31 �C ' �@�����Y�b����
Public Const RowTask12SmallTransferHundred   As Integer = 32    ' �����b�� 32 �C ' �p��ʤ����Y�b����
Public Const RowTask12SmallTransferMany      As Integer = 33    ' �����b�� 33 �C ' �����Y�b����

' Note: this one has to be the last as the item count is not fixed
' 3.2 ������ �������ȩҰ_�l���C���A�Ъ`�N�A�p�G�W�� 2.1 ���C�O���֬�Ĳ�쪺�ɭԡA�o�ӭȬO�i�H���U�վ㪺�C
Public Const RowTopCounterparty As Integer = 60                 ' �����b�� 60 �C ' �q�o�C�H�U�A�O3.2����
' =========================================================