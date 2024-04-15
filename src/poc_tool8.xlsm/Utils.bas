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

Public Function ExtractNumbersPrefix(inputString As String) As String
    Dim i      As Integer
    Dim result As String

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

Public Function SumDictValues(dict As Object) As Double
    Dim key As Variant
    Dim total As Double

    ' Iterate through each item in the dictionary
    For Each key In dict.keys
        ' Add the value of the current item to the total
        If VarType(dict(key)) = vbDouble Then
            total = total + dict(key)
        End If
    Next key

    ' Return the total sum
    SumDictValues = total
End Function

Public Function RemoveLeadingZeros(ByVal str) As String
    Dim tmpStr As String
    tmpStr = str
    While (Left(tmpStr, 1) = "0") AND (tmpStr <> "")
        tmpStr = Right(tmpStr, Len(tmpStr)-1)
    Wend

    RemoveLeadingZeros = tmpStr
End Function

Public Function GuessIfIsTheSame(ByVal str1, ByVal str2) As Boolean
    Dim bIsTheSame As Boolean
    bIsTheSame = True
    tmpStr1 = str1
    tmpStr2 = str2
    If Len(tmpStr1) <> Len(tmpStr2) Then
        bIsTheSame = False
    Else
        While ((tmpStr1 <> "") AND (tmpStr2 <> ""))
            If (Left(tmpStr1, 1) <> "*") AND (Left(tmpStr2, 1) <> "*") AND (Left(tmpStr1, 1) <> Left(tmpStr2, 1)) Then
                bIsTheSame = False
                GoTo EndOfFunc
            Else
                tmpStr1 = Right(tmpStr1, Len(tmpStr1)-1)
                tmpStr2 = Right(tmpStr2, Len(tmpStr2)-1)
            End If
        Wend
    End If

EndOfFunc:
    GuessIfIsTheSame = bIsTheSame
End Function

Public Function SkipLeadingZeros(account As String) As String
    Dim skipzerodAcc As String
    skipzeroAcc = Trim(account)
    
    Do While Left(skipzeroAcc, 1) = "0" And Len(skipzeroAcc) > 1
        skipzeroAcc = Mid(skipzeroAcc, 2)
    Loop
    
    SkipLeadingZeros = skipzeroAcc
End Function