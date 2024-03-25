Attribute VB_Name = "Globals"
Option Explicit

Public UpdateSimplePage As Boolean

Public DictBlacklist  As Object
Public DictVirtualAcc As Object

' Main Account Name & Account ID
Public MainAccName As String
Public MainAccId   As String

' 用來紀錄是否有 ATM，來開關 ATM 相關的 PivotTable
Public GotATM As Long

Public HeaderPivotTable3 As Range
Public HeaderPivotTable4 As Range
Public HeaderPivotTable5 As Range
Public HeaderPivotTable6 As Range
