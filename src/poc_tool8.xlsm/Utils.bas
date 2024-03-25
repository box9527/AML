Attribute VB_Name = "Utils"

' Get array length
Public Function GetLength(A As Variant) As Integer
    If IsEmpty(A) Then
        GetLength = 0
    Else
        GetLength = UBound(A) - LBound(A) + 1
    End If
End Function

' Check if a dict contains an item with item name
Public Function ObjectContainsItem(ByRef items As Object, itemToFind As Variant) As Boolean
    Dim bExisted As Boolean
    Dim i As Long

    ' Iterate through the array
    For i = 1 To items.Count
        If items(i).name = itemToFind Then
            bExisted = True
            Exit For
        End If
    Next i

    ' Item not found
    ObjectContainsItem = bExisted
End Function

Public Function ArrayContainsItem(ByRef arr As Variant, itemToFind As String) As Boolean
    Dim bExisted As Boolean
    Dim i As Long

    For i = LBound(arr) To UBound(arr)
        If arr(i) = itemToFind Then
            bExisted = True
            Exit For
        End If
    Next i

    ArrayContainsItem = bExisted
End Function

Public Sub CountOne(ByRef Num As Variant)
    Num = Num + 1
End Sub

Public Sub NormalizeCellStartEnd(ByRef cellStart As String, ByRef cellEnd As String)
    If Len(cellStart) > 0 And Len(cellEnd) = 0 Then
        cellEnd = cellStart
    End If
    If Len(cellStart) = 0 And Len(cellEnd) > 0 Then
        cellStart = cellEnd
    End If
End Sub

Public Sub NormalizeCellRowCol(ByRef cellRow As Long, ByRef cellCol As Long)
    If cellRow > 0 And cellCol <= 0 Then
        cellCol = cellRow
    End If
    If cellRow <= 0 And cellCol > 0 Then
        cellRow = cellCol
    End If
End Sub

Public Sub NormalizeDataFormat(ByRef dataFormat As String)
    If Len(dataFormat) < 0 Or ((dataFormat <> DateFormat) And _
    (dataFormat <> TimeFormat) And (dataFormat <> NumberFormat) And _
    (dataFormat <> ForceStringFormat) And (dataFormat <> GeneralFormat)) Then
        dataFormat = ForceStringFormat
    End If
End Sub

Public Sub NormalizeDataHDirection(ByRef direction As Integer)
    If direction <> xlLeft Or direction <> xlRight Or direction <> xlCenter Then
        direction = xlCenter
    End If
End Sub

Public Sub NormalizeFontSize(ByRef size As Integer)
    If size > (FontSize * 4) Or size < FontSize Then
        size = FontSize
    End If
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

Public Function IsArrayEmpty(arr As Variant) As Boolean
    If UBound(arr) = 1 And arr(1) = EmptyArrayValue Then
        IsArrayEmpty = True
    Else
        IsArrayEmpty = False
    End If
End Function
