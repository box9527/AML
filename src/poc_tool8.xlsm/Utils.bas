Attribute VB_Name = "Utils"

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

Public Sub QuickSort(ByRef arrValues() As Variant, ByRef arrKeys() As Variant, ByVal low As Long, ByVal high As Long)
    Dim pivot       As Variant
    Dim i           As Long
    Dim j           As Long
    Dim tempValue   As Variant
    Dim tempKey     As Variant
    i = low
    j = high
    pivot = arrValues((low + high) \ 2)
    Do While i <= j
        Do While arrValues(i) > pivot
            i = i + 1
        Loop
        Do While arrValues(j) < pivot
            j = j - 1
        Loop
        If i <= j Then
            ' Swap values
            tempValue = arrValues(i)
            arrValues(i) = arrValues(j)
            arrValues(j) = tempValue
            ' Swap keys
            tempKey = arrKeys(i)
            arrKeys(i) = arrKeys(j)
            arrKeys(j) = tempKey
            i = i + 1
            j = j - 1
        End If
    Loop
    If low < j Then QuickSort arrValues, arrKeys, low, j
    If i < high Then QuickSort arrValues, arrKeys, i, high
End Sub

Public Function ContainsNumbers(inputString As String) As Boolean
    Dim i           As Integer
    Dim charCode    As Integer
    Dim hasNumbers  As Boolean
    ' Initialize variables
    hasNumbers = False
    ' Loop through each character in the string
    For i = 1 To Len(inputString)
        ' Check if the character is a number (ASCII code between 48 and 57)
        charCode = Asc(Mid(inputString, i, 1))
        If charCode >= 48 And charCode <= 57 Then
            hasNumbers = True
        Else
            hasNumbers = False
            Exit For
        End If
    Next i
    ' Return the result
    ContainsNumbers = hasNumbers
End Function
