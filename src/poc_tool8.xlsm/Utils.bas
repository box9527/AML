Attribute VB_Name = "Utils"
Public Function CheckPivotItemExisted(objPivotTable As PivotTable, strPivotFieldName As String, strPivotItemName As String) As Boolean
    '==========================
    ' �o�̥i��|�]���ѪR�X�Ӫ�����K�n���]�t"03�����b"�ӥ��ѡA
    ' �[�W�@�� For loop Check�A�ˬd�즳item�A�N���}�C
    '.PivotFields("����K�n").CurrentPage = "03�����b"
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
