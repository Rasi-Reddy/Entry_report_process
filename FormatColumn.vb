Sub FormatColumns(colNames As Variant, startRow As Integer)

    Dim col As Variant
    Dim lastRow As Long
    
    ' Find last row (based on Q or entire sheet)
    lastRow = Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    
    ' Check if single column or multiple
    If IsArray(colNames) Then
        
        ' Multiple columns
        For Each col In colNames
            With Range(col & startRow & ":" & col & lastRow)
                .Value = .Value

.NumberFormat = "[$R$-pt-BR] #,##0;[$R$-pt-BR] -#,##0"
                
            End With
        Next col
        
    Else
        
        ' Single column
        With Range(colNames & startRow & ":" & colNames & lastRow)
            .Value = .Value

.NumberFormat = "[$R$-pt-BR] #,##0;[$R$-pt-BR] -#,##0"
           
        End With
        
    End If

End Sub