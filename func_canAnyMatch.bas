Attribute VB_Name = "func_canAnyMatch"
Public Function canAnyMatch(board As Range) As Boolean

canAnyMatch = False

Dim c As Range

For Each c In board
    Select Case c.row()
        Case 1:
            If legalMatch(board, c, c.Offset(1, 0)) Then
                canAnyMatch = True
                Exit Function
            End If
        Case 2 To 9:
            If legalMatch(board, c, c.Offset(1, 0)) Then
                canAnyMatch = True
                Exit Function
            End If
            If legalMatch(board, c, c.Offset(-1, 0)) Then
                canAnyMatch = True
                Exit Function
            End If
        Case 10:
            If legalMatch(board, c, c.Offset(-1, 0)) Then
                canAnyMatch = True
                Exit Function
            End If
    End Select
    
    Select Case c.Column()
        Case 1:
            If legalMatch(board, c, c.Offset(0, 1)) Then
                canAnyMatch = True
                Exit Function
            End If
        Case 2 To 9:
            If legalMatch(board, c, c.Offset(0, 1)) Then
                canAnyMatch = True
                Exit Function
            End If
            If legalMatch(board, c, c.Offset(0, -1)) Then
                canAnyMatch = True
                Exit Function
            End If
        Case 10:
            If legalMatch(board, c, c.Offset(0, -1)) Then
                canAnyMatch = True
                Exit Function
            End If
    End Select
Next c

End Function
