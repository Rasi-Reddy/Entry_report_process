Sub FormatColumns(colNames As Variant, startRow As Integer)

    Dim lastRow As Long
    Dim col As Variant
    
    lastRow = Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    
    ' Force array handling
    If IsArray(colNames) Then
        For Each col In colNames
            Range(col & startRow & ":" & col & lastRow).NumberFormat = "[$R$-pt-BR] #,##0;[$R$-pt-BR] -#,##0"
        Next col
    Else
        ' fallback if single column passed
        Range(colNames & startRow & ":" & colNames & lastRow).NumberFormat = "[$R$-pt-BR] #,##0;[$R$-pt-BR] -#,##0"
    End If

End Sub