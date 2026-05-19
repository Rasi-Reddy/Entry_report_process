Sub FormatColumns(colNames As Variant, startRow As Integer)

    Dim lastRow As Long
    Dim col As Variant
    
    lastRow = Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    
    If IsArray(colNames) Then
        For Each col In colNames
            Range(col & startRow & ":" & col & lastRow).NumberFormat = "#,##0;-#,##0"
        Next col
    Else
        Range(colNames & startRow & ":" & colNames & lastRow).NumberFormat = "#,##0;-#,##0"
    End If

End Sub