Attribute VB_Name = "func_legalMatch"
Public Function legalMatch(board As Range, Optional cellA As Range, Optional cellB As Range) As Boolean

legalMatch = False

Dim cells(1 To 2) As Range


If cellA Is Nothing Then
    Set cells(1) = Selection(1)
    Set cells(2) = Selection(2)
Else
    Set cells(1) = cellA
    Set cells(2) = cellB
End If

'Dim tempVal As Variant
'
'tempVal = cells(2).Value
'cells(2).Value = cells(1).Value
'cells(1).Value = tempVal

Dim cellsValue(1 To 2) As String
cellsValue(1) = cells(2).Value
cellsValue(2) = cells(1).Value


Dim row As Integer
Dim col As Integer
Dim i As Integer

        For i = 1 To 2
        
            row = cells(i).row()
            col = cells(i).Column()
                    
            Select Case row
                Case 1:
                    If cellsValue(i) = cells(i).Offset(1, 0).Value And _
                        cellsValue(i) = cells(i).Offset(2, 0).Value Then
                        legalMatch = True
                        GoTo TheEnd
                    End If
                Case 2:
                    If cellsValue(i) = cells(i).Offset(1, 0).Value And _
                        cellsValue(i) = cells(i).Offset(2, 0).Value Then
                        legalMatch = True
                        GoTo TheEnd
                    End If
                    If cellsValue(i) = cells(i).Offset(-1, 0).Value And _
                        cellsValue(i) = cells(i).Offset(1, 0).Value Then
                        legalMatch = True
                        GoTo TheEnd
                    End If
                Case 3 To 8:
                    If cellsValue(i) = cells(i).Offset(1, 0).Value And _
                        cellsValue(i) = cells(i).Offset(2, 0).Value Then
                        legalMatch = True
                        GoTo TheEnd
                    End If
                    If cellsValue(i) = cells(i).Offset(-1, 0).Value And _
                        cellsValue(i) = cells(i).Offset(1, 0).Value Then
                        legalMatch = True
                        GoTo TheEnd
                    End If
                    If cellsValue(i) = cells(i).Offset(-1, 0).Value And _
                        cellsValue(i) = cells(i).Offset(-2, 0).Value Then
                        legalMatch = True
                        GoTo TheEnd
                    End If
                Case 9:
                    If cellsValue(i) = cells(i).Offset(-1, 0).Value And _
                        cellsValue(i) = cells(i).Offset(1, 0).Value Then
                        legalMatch = True
                        GoTo TheEnd
                    End If
                    If cellsValue(i) = cells(i).Offset(-1, 0).Value And _
                        cellsValue(i) = cells(i).Offset(-2, 0).Value Then
                        legalMatch = True
                        GoTo TheEnd
                    End If
                Case 10:
                    If cellsValue(i) = cells(i).Offset(-1, 0).Value And _
                        cellsValue(i) = cells(i).Offset(-2, 0).Value Then
                        legalMatch = True
                        GoTo TheEnd
                    End If
            End Select
        
            Select Case col
                Case 1:
                    If cellsValue(i) = cells(i).Offset(0, 1).Value And _
                        cellsValue(i) = cells(i).Offset(0, 2).Value Then
                        legalMatch = True
                        GoTo TheEnd
                    End If
                Case 2:
                    If cellsValue(i) = cells(i).Offset(0, 1).Value And _
                        cellsValue(i) = cells(i).Offset(0, 2).Value Then
                        legalMatch = True
                        GoTo TheEnd
                    End If
                    If cellsValue(i) = cells(i).Offset(0, -1).Value And _
                        cellsValue(i) = cells(i).Offset(0, 1).Value Then
                        legalMatch = True
                        GoTo TheEnd
                    End If
                Case 3 To 8:
                    If cellsValue(i) = cells(i).Offset(0, 1).Value And _
                        cellsValue(i) = cells(i).Offset(0, 2).Value Then
                        legalMatch = True
                        GoTo TheEnd
                    End If
                    If cellsValue(i) = cells(i).Offset(0, -1).Value And _
                        cellsValue(i) = cells(i).Offset(0, 1).Value Then
                        legalMatch = True
                        GoTo TheEnd
                    End If
                    If cellsValue(i) = cells(i).Offset(0, -1).Value And _
                        cellsValue(i) = cells(i).Offset(0, -2).Value Then
                        legalMatch = True
                        GoTo TheEnd
                    End If
                Case 9:
                    If cellsValue(i) = cells(i).Offset(0, -1).Value And _
                        cellsValue(i) = cells(i).Offset(0, 1).Value Then
                        legalMatch = True
                        GoTo TheEnd
                    End If
                    If cellsValue(i) = cells(i).Offset(0, -1).Value And _
                        cellsValue(i) = cells(i).Offset(0, -2).Value Then
                        legalMatch = True
                        GoTo TheEnd
                    End If
                Case 10:
                    If cellsValue(i) = cells(i).Offset(0, -1).Value And _
                        cellsValue(i) = cells(i).Offset(0, -2).Value Then
                        legalMatch = True
                        GoTo TheEnd
                    End If
            End Select
        
        Next i
    
TheEnd:

'tempVal = cells(2).Value
'cells(2).Value = cells(1).Value
'cells(1).Value = tempVal

End Function


