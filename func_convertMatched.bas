Attribute VB_Name = "func_convertMatched"
Public Function convertMatched(board As Range) As Integer

convertMatched = 0

Dim col As Integer
Dim row As Integer


For row = 1 To 8
    For col = 1 To 10
    
        If CStr(board(row, col)) = CStr(board(row + 1, col)) Or _
                    board(row, col) = matchCompare(CStr(board(row + 1, col))) Then
            If CStr(board(row, col)) = CStr(board(row + 2, col)) Or _
                    CStr(board(row, col)) = matchCompare(CStr(board(row + 2, col))) Then
                board(row, col) = matchChange(board(row, col))
                board(row + 1, col) = matchChange(board(row + 1, col))
                board(row + 2, col) = matchChange(board(row + 2, col))
                convertMatched = convertMatched + 1
            End If
        End If
    
    Next col
Next row

For col = 1 To 8
    For row = 1 To 10
    
        If CStr(board(row, col)) = CStr(board(row, col + 1)) Or _
                    board(row, col) = matchCompare(CStr(board(row, col + 1))) Then
            If CStr(board(row, col)) = CStr(board(row, col + 2)) Or _
                    CStr(board(row, col)) = matchCompare(CStr(board(row, col + 2))) Then
                board(row, col) = matchChange(board(row, col))
                board(row, col + 1) = matchChange(board(row, col + 1))
                board(row, col + 2) = matchChange(board(row, col + 2))
                convertMatched = convertMatched + 1
            End If
        End If
    
    Next row
Next col


End Function


