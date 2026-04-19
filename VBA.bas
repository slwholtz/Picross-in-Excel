Attribute VB_Name = "Module1"
Sub PUZZLEIZE()
    Dim rows As Integer, cols As Integer
    rows = InputBox("How many rows?")
    cols = InputBox("How many columns?")

    Dim maxCountCols As Integer
    maxCountCols = 0
    Dim lsts() As Collection
    ReDim lsts(1 To cols)
    
    For j = 1 To cols
        Set lsts(j) = New Collection
        For i = 1 To rows
            If Cells(i, j).Interior.Color = RGB(0, 0, 0) Then
                Dim Count As Integer
                Count = 0
                While Cells(i, j).Interior.Color = RGB(0, 0, 0)
                    Count = Count + 1
                    i = i + 1
                Wend
                lsts(j).Add (Count)
                If lsts(j).Count > maxCountCols Then
                    maxCountCols = lsts(j).Count
                End If
            End If
        Next i
    Next j

    Dim lsts2() As Collection
    ReDim lsts2(1 To rows)
    Dim maxCountRows As Integer
    maxCountRows = 0
    For i = 1 To rows
        Set lsts2(i) = New Collection
        For j = 1 To cols
            If Cells(i, j).Interior.Color = RGB(0, 0, 0) Then
                Count = 0
                While Cells(i, j).Interior.Color = RGB(0, 0, 0)
                    Count = Count + 1
                    j = j + 1
                Wend
                lsts2(i).Add (Count)
                If lsts2(i).Count > maxCountRows Then
                    maxCountRows = lsts2(i).Count
                End If
            End If
        Next j
    Next i
    
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets.Add
    ws.Cells.RowHeight = 22.5
    ws.Cells.ColumnWidth = 3.57
    ws.Range("A1").Value = rows & " x " & cols
    
    For j = 1 To cols
        ' add case for empty
        For i = 1 To lsts(j).Count
            ws.Cells(maxCountCols - i + 1, maxCountRows + j) = lsts(j)(i)
        Next i
    Next j
    
    For i = 1 To rows
        'add case for empty
        For j = 1 To lsts2(i).Count
            ws.Cells(maxCountCols + i, maxCountRows - j + 1) = lsts2(i)(j)
        Next j
    Next i
End Sub
