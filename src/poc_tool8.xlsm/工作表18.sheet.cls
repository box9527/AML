
Private Sub Worksheet_PivotTableUpdate(ByVal Target As PivotTable)
     Exit Sub ' FIXME: to let user change filtering option. Need to reserve enough space for pivot table
    
    If Target.name = "MyPivotTable3" Or Target.name = "MyPivotTable4" Or Target.name = "MyPivotTable5" Or Target.name = "MyPivotTable6" Then
        On Error Resume Next
        If Not pivotRange Is Nothing Then
            With Target
              '   .PivotFields("轉出入帳號").AutoSort xlDescending, "計數 - In"
                 
            End With
        End If
        On Error GoTo 0
    End If


   If Target.name = "MyPivotTable3" Or Target.name = "MyPivotTable4" Or Target.name = "MyPivotTable5" Or Target.name = "MyPivotTable6" Then
        On Error Resume Next
        Set pivotRange = Target.TableRange1
        If Not pivotRange Is Nothing Then
            With Target.TableRange1
                .Rows.Hidden = False
                If .Rows.Count > 6 Then
                       .Rows("6:" & CStr(.Rows.Count - 1)).Hidden = True
                End If
                
            End With
        End If
        On Error GoTo 0
    End If
End Sub