Attribute VB_Name = "Utils"
Public Function CheckPivotItemExisted(objPivotTable As PivotTable, strPivotFieldName As String, strPivotItemName As String) As Boolean
    '==========================
    ' 這裡可能會因為解析出來的交易摘要不包含"03跨行轉帳"而失敗，
    ' 加上一個 For loop Check，檢查到有item，就離開。
    '.PivotFields("交易摘要").CurrentPage = "03跨行轉帳"
    Dim bExisted As Boolean
    For Each pivot_item In objPivotTable.PivotFields(strPivotFieldName).PivotItems
        If pivot_item.name = strPivotItemName Then
            bExisted = True
            Exit For
            Exit Function
        End If
    Next pivot_item

    Debug.Print bExisted
    bExisted = False
    Exit Function
    
    '==========================
End Function

' Get array length
Public Function GetLength(A As Variant) As Integer
    If IsEmpty(A) Then
        GetLength = 0
    Else
        GetLength = UBound(A) - LBound(A) + 1
    End If
End Function

' Check if a array contains an item with item name
Public Function ObjectContainsItem(ByRef items As Object, itemToFind As Variant) As Boolean
    Dim bExisted As Boolean
    Dim i As Long
    
    ' Iterate through the array
    For i = 1 To items.Count
        If items(i).name = itemToFind Then
            bExisted = True
            Exit For
            Exit Function
        End If
    Next i
    
    ' Item not found
    bExisted = False
    Exit Function
End Function

Public Sub CountOne(ByRef Num As Integer)
    Num = Num + 1
End Sub

Public Function IsItemInArray(StrItem As String) As Boolean
    'Dim IsItemInArray As Boolean
    ' Version control files
    Dim VerCtrlFiles(VerCtrlFilesSize - 1) As String
    VerCtrlFiles(0) = "Build"
    VerCtrlFiles(1) = "ErrorHandling"
    VerCtrlFiles(2) = "Formatter"
    VerCtrlFiles(3) = "NamedRanges"
    VerCtrlFiles(4) = "Menu"
    VerCtrlFiles(5) = "Test"
    VerCtrlFiles(6) = "XMLexporter"
    VerCtrlFiles(7) = "CustomActions"
    VerCtrlFiles(8) = "EventListener"
    VerCtrlFiles(9) = "MyCustomActions"

    Dim i As Integer
    For i = LBound(VerCtrlFiles) To UBound(VerCtrlFiles)
        If StrComp(VerCtrlFiles(i), StrItem, vbTextCompare) = 0 Then
            IsItemInArray = True
            Exit Function
        End If
    Next i
    IsItemInArray = False
End Function

Sub TestIsItemInArray()
    Dim ItemToCheck As String
    ItemToCheck = "Build"
    
    If IsItemInArray(ItemToCheck) Then
        MsgBox ItemToCheck & " exists in the array."
    Else
        MsgBox ItemToCheck & " does not exist in the array."
    End If
End Sub

