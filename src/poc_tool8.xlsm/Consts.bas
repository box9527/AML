Attribute VB_Name = "Consts"
Option Explicit

' Derivative watch-listed account
Public Const FileBadAcc     As String = "ĵ�ܤ�.xlsx"
Public Const SheetNameBadAcc As String = "ĵ�ܤ�"

' Virtual account
Public Const FileVirtualAcc     As String = "�����b��.xlsx"
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
Public Const ColorYellow2    As Long = &H86FFFE


' Sheet name definition
Public Const SheetNameOrginal      As String = "1��l���"
Public Const SheetNameMain         As String = "2.1�D����"
Public Const SheetNameInputData    As String = "2.2�M�����"
Public Const SheetNameSimple       As String = "3.1�������"
Public Const SheetNameMoney        As String = "3.2���y�P������"
Public Const SheetNameBranch       As String = "����M��"
Public Const SheetNameIntermediate As String = "�Ȧs��"
Public Const SheetNameLabel As String = "�ۭq�Хܳ]�w"
Public Const RowDataBegin       As Integer = 9

' Total count of version control files
Public Const VerCtrlFilesSize As Integer = 10



