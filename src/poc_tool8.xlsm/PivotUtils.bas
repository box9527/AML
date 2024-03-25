Attribute VB_Name = "PivotUtils"

Public Function CreateCaches(dataRange As Range) As PivotCache
    Dim cache As pivotCache
    If Not dataRange Is Nothing Then
        Set cache = ThisWorkbook.PivotCaches.Create( _
                               SourceType:=xlDatabase, _
                               SourceData:=dataRange)
    End If

    Set CreateCaches = cache
End Function

Public Function CheckPivotItemExisted(ByRef ws As Worksheet, strPivotTableName As String, _
                                      strPivotFieldName As String, strPivotItemName As String) As Boolean
    Dim bExisted As Boolean
    If Not ws Is Nothing Then
        With ws.PivotTables(strPivotTableName).PivotFields(strPivotFieldName)
            For Each pivot_item In .PivotItems
                If pivot_item.name = strPivotItemName Then
                    bExisted = True
                    Exit For
                End If
            Next pivot_item
        End With
    End If

    CheckPivotItemExisted = bExisted
End Function
