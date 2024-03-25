Attribute VB_Name = "ExtraData"
Option Explicit

Private isDictInited As Boolean

Sub InitDicts()
    If isDictInited = True Then
        Debug.Print "dict inited before!!"
        Exit Sub
    End If

    Set DictBlacklist = GetBlacklist()

    Set DictVirtualAcc = GetVirtualAcc()

    ' disable this to load them every thime
    'isDictInited = True

End Sub

Public Function CheckVAccount(bankID As String, account As String) As String
    Dim ret         As String
    Dim tmp         As String
    Dim b           As Boolean

    'Debug.Print "Check " & bankID & " " & account
    ret = ""
    tmp = ""

    If Len(bankID) = 0 Or Len(account) = 0 Then
        ret = ""

    ElseIf Not DictVirtualAcc.Exists(bankID) Then
        ret = ""

    Else
        Dim key     As Variant
        For Each key In DictVirtualAcc(bankID)
            'Debug.Print "Got key  " & key & " value: " & dictVirtualAcc(bankID)(key)
            b = ContainPrefix(CStr(key), account)
            If b = True Then
                ret = ret & " " & DictVirtualAcc(bankID)(key)
            End If

        Next key
    End If

    CheckVAccount = Trim(ret)

End Function

Function ContainPrefix(Prefix As String, Text As String) As Boolean
    ' Check if Text contains Prefix at the beginning
    ContainPrefix = (InStr(1, Text, Prefix, vbTextCompare) = 1)
End Function

Private Function GetVirtualAcc() As Object
    Dim wsRef       As Worksheet
    Dim dict        As Object
    Dim lastRow     As Long
    Dim i           As Long
    Dim wbRef       As Workbook
    Dim currentPath As String
    Dim isOpened    As Boolean

    isOpened = True
    ' Get the current working directory
    currentPath = ThisWorkbook.Path

    On Error Resume Next
    Set wbRef = Workbooks(FileVirtualAcc)

    On Error GoTo 0
    If wbRef Is Nothing Then
        ' Workbook is not open, open it
        isOpened = False
        Set wbRef = Workbooks.Open(currentPath & "\" & FileVirtualAcc)

    End If

    ' Set the reference worksheet
    Set wsRef = wbRef.Sheets(SheetNameVirtualAcc)

    Dim bank        As String
    Dim strRule     As String
    Dim prefixRule  As String
    Dim colKey      As String
    Dim colVal      As String
    Dim colUsage    As String

    colKey = "A"
    colVal = "B"
    colUsage = "E"

    Set dict = CreateObject("Scripting.Dictionary")

    lastRow = wsRef.Cells(wsRef.Rows.Count, colKey).End(xlUp).row
    For i = 2 To lastRow
        bank = CStr(wsRef.Cells(i, colKey).value)
        If Not IsEmpty(bank) Then
            strRule = wsRef.Cells(i, colVal).value & " (" & bank & " " & wsRef.Cells(i, colUsage).value & ")"

            prefixRule = ExtractNumbersPrefix(strRule)
            If Not dict.Exists(bank) Then
                Dim nestedDict1 As Object
                Set nestedDict1 = CreateObject("Scripting.Dictionary")

                nestedDict1.Add prefixRule, strRule
                dict.Add bank, nestedDict1
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

Function ExtractNumbersPrefix(inputString As String) As String
    Dim i           As Integer
    Dim result      As String

    ' Loop through each character in the input string
    For i = 1 To Len(inputString)
        ' Check if the character is a number
        If IsNumeric(Mid(inputString, i, 1)) Then
            ' Append the numeric character to the result
            result = result & Mid(inputString, i, 1)
        Else

            Exit For

        End If
    Next i

    ' Return the result containing numbers only
    ExtractNumbersPrefix = result
End Function
