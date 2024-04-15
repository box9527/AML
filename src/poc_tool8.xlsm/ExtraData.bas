Attribute VB_Name = "ExtraData"
Option Explicit

Sub InitDicts()
    If IsDictInited = True Then
        Debug.Print "dict inited before!!"
        Exit Sub
    End If

    Set DictBlacklist = GetBlacklist()

    Set DictVirtualAcc = GetVirtualAcc()

    ' disable this to load them every thime
    IsDictInited = True
End Sub

Public Function ConvVAccName(bankID As String, account As String) As String
    Dim ret As String
    Dim b   As Boolean

    account = Utils.SkipLeadingZeros(account)
    If (Len(bankID) > 0) And (Len(account) > 0) And (DictVirtualAcc.Exists(bankID) = True) Then
        Dim key     As Variant
        For Each key In DictVirtualAcc(bankID)
            'Debug.Print "Got key  " & key & " value: " & dictVirtualAcc(bankID)(key)
            b = ContainPrefix(CStr(key), account)
            If b = True Then
                ret = ret & " " & DictVirtualAcc(bankID)(key)
            End If

        Next key
    End If

    ConvVAccName = CStr(Trim(ret))
End Function

Public Function IsWarningAcc(account As String) As Boolean
    Dim isWarn As Boolean
    If DictBlacklist.Exists(account) = True Then
        isWarn = True
    End If

    IsWarningAcc = isWarn
End Function

Function ContainPrefix(Prefix As String, Text As String) As Boolean
    ' Check if Text contains Prefix at the beginning
    ContainPrefix = (InStr(1, Text, Prefix, vbTextCompare) = 1)
End Function

Private Function GetVirtualAcc() As Object
    Dim wsRef    As Worksheet
    Dim lastRow  As Long
    Dim i        As Long
    Dim wbRef    As Workbook
    Dim currPath As String
    Dim isOpened As Boolean

    isOpened = True
    ' Get the current working directory
    currPath = ThisWorkbook.Path

    On Error Resume Next
    Set wbRef = Workbooks(FileVirtualAcc)
    On Error GoTo 0

    If wbRef Is Nothing Then
        ' Workbook is not open, open it
        isOpened = False
        Set wbRef = Workbooks.Open(currPath & "\" & FileVirtualAcc)
    End If

    ' Set the reference worksheet
    Set wsRef = wbRef.Sheets(SheetNameVirtualAcc)

    Dim bank       As String
    Dim accN       As String
    Dim strRule    As String
    Dim prefixRule As String
    Dim colKey     As String
    Dim colVal     As String
    Dim colAcc     As String
    Dim colUsage   As String

    colKey = "A" ' 行庫別 ' bank
    colVal = "B" ' 帳號
    colAcc = "C" ' 戶名
    colUsage = "E" ' 用途: 公司戶、虛擬帳號...
    ' D: 統編、F: 備註

    ' 這裡必須要用這種創建 dictionary 的方式做，
    ' 用 As New Scripting.Dictionary 會遇到 "鍵值重複使用的問題"
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    lastRow = wsRef.Cells(wsRef.Rows.Count, colKey).End(xlUp).row

    For i = 2 To lastRow
        bank = CStr(wsRef.Cells(i, colKey).value)
        accN = CStr(wsRef.Cells(i, colAcc).value)

        If Not IsEmpty(bank) Then
            strRule = wsRef.Cells(i, colVal).value & " (" & bank & " " & wsRef.Cells(i, colUsage).value & ")" & _
                      GeneralDelimiter & accN
            prefixRule = Utils.ExtractNumbersPrefix(strRule)

            If Not dict.Exists(bank) Then
                'Dim nestedDict As New Scripting.Dictionary
                Dim nestedDict As Object
                Set nestedDict = CreateObject("Scripting.Dictionary")

                ' dict.Add key, item
                nestedDict.Add prefixRule, strRule
                dict.Add bank, nestedDict
            Else
                'Debug.Print "in bank " & bank & " add a New rule: " & prefixRule & "    " & strRule
                dict(bank).Add prefixRule, strRule
            End If
        End If
    Next i

    If isOpened = False And wbRef.name = FileVirtualAcc Then
        wbRef.Close SaveChanges:=False
    End If

    Set GetVirtualAcc = dict
End Function

Private Function GetBlacklist() As Object
    Dim wsRef       As Worksheet
    Dim dict        As Object
    Dim lastRow     As Long
    Dim i           As Long
    Dim blacklistItem As Variant
    Dim wbRef       As Workbook
    Dim currentPath As String
    Dim isOpened    As Boolean

    isOpened = True
    ' Get the current working directory
    currentPath = ThisWorkbook.Path
    On Error Resume Next
    Set wbRef = Workbooks(FileBadAcc)
    On Error GoTo 0
    If wbRef Is Nothing Then
        ' Workbook is not open, open it
        isOpened = False
        Set wbRef = Workbooks.Open(currentPath & "\" & FileBadAcc)
    End If

    ' Set the reference worksheet
    Set wsRef = wbRef.Sheets(SheetNameBadAcc)
    Set dict = ExtraData.GetKVPairsToDict(wsRef, "G", "G")

    If isOpened = False And wbRef.name = FileBadAcc Then
        wbRef.Close SaveChanges:=False
    End If

    ' Now you have the items in the "BLACKLIST" sheet in the dictionary
    ' You can use the dictionary as needed
    Set GetBlacklist = dict

End Function

Public Function GetKVPairsToDict(wsRef As Worksheet, colKey As String, colVal As String) As Object
    Dim dict        As Object
    Dim lastRow     As Long
    Dim i           As Long
    Dim item        As Variant

    Set dict = CreateObject("Scripting.Dictionary")

    lastRow = wsRef.Cells(wsRef.Rows.Count, colKey).End(xlUp).row
    ' Loop through column G and add items to the dictionary
    For i = 1 To lastRow
        item = CStr(wsRef.Cells(i, colKey).value)
        If Not IsEmpty(item) Then
            If Not dict.Exists(item) Then
                'Debug.Print "add blacklist item: " & blacklistItem
                dict.Add item, wsRef.Cells(i, colVal).value
            End If
        End If
    Next i

    ' You can use the dictionary as needed
    Set GetKVPairsToDict = dict

End Function
