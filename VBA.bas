Attribute VB_Name = "Module1"
Sub PUZZLEIZE()
    ' Fill in the top part of the puzzle
    For j = 8 To 21
        Dim lst As New Collection
        For i = 10 To 25
            If Cells(i, j).Interior.Color = RGB(0, 0, 0) Then
                Dim Count As Integer
                Count = 0
                While Cells(i, j).Interior.Color = RGB(0, 0, 0)
                    Count = Count + 1
                    i = i + 1
                Wend
                lst.Add (Count)
            End If
        Next i
        For i = 1 To lst.Count
            Cells(9 - i + 1, j).Value = lst(lst.Count - i + 1)
        Next i
        Set lst = Nothing
    Next j
    'Bottom Part
    
    For i = 10 To 25
        Set lst = New Collection
        For j = 8 To 21
            If Cells(i, j).Interior.Color = RGB(0, 0, 0) Then
                Count = 0
                While Cells(i, j).Interior.Color = RGB(0, 0, 0)
                    Count = Count + 1
                    j = j + 1
                Wend
                lst.Add (Count)
            End If
        Next j
        For j = 1 To lst.Count
            Cells(i, 7 - j + 1).Value = lst(lst.Count - j + 1)
        Next j
        Set lst = Nothing
    Next i
End Sub
