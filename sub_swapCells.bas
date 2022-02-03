Attribute VB_Name = "sub_swapCells"
Public Sub swapCells()

ThisWorkbook.ActiveSheet.Range("P13").Value = ""
Dim tempvalue As String

Dim wb As Workbook
Set wb = ThisWorkbook

Dim ws As Worksheet
Set ws = wb.Sheets("Board")

Dim board As Range
Set board = ws.Range("A1:J10")

Dim gemTracker(1 To 7) As Integer

If TypeName(Selection) = "Range" Then
    If Selection.Count = 2 Then
        If legalMatch(board) Then
            tempvalue = Selection(1).Value
            Selection(1).Value = Selection(2).Value
            Selection(2).Value = tempvalue
        Else
            ThisWorkbook.ActiveSheet.Range("P13").Value = "Will Not Match :("
            Exit Sub
        End If
    Else
        ThisWorkbook.ActiveSheet.Range("P13").Value = "Wrong Selection Size."
    End If
Else
    'pass
End If

Dim activity As Integer

Do
    activity = updateBoard(board, gemTracker)
Loop Until activity = 0

If canAnyMatch(board) = False Then
        ThisWorkbook.ActiveSheet.Range("P13").Value = "Game Over!"
End If

'update history: gems
ws.Range("T10") = ws.Range("T9")
ws.Range("T9") = ws.Range("T8")
ws.Range("T8") = ws.Range("T7")
ws.Range("T7") = ws.Range("T6")
ws.Range("T6") = 0

'update history: Matches
ws.Range("U10") = ws.Range("U9")
ws.Range("U9") = ws.Range("U8")
ws.Range("U8") = ws.Range("U7")
ws.Range("U7") = ws.Range("U6")
ws.Range("U6") = 0

'update history: Multipliers
ws.Range("V10") = ws.Range("V9")
ws.Range("V9") = ws.Range("V8")
ws.Range("V8") = ws.Range("V7")
ws.Range("V7") = ws.Range("V6")
ws.Range("V6") = 0

'update history: Score
ws.Range("W10") = ws.Range("W9")
ws.Range("W9") = ws.Range("W8")
ws.Range("W8") = ws.Range("W7")
ws.Range("W7") = ws.Range("W6")
ws.Range("W6") = 0

'update current gem score
ws.Range("Q4") = gemTracker(1)
ws.Range("Q5") = gemTracker(2)
ws.Range("Q6") = gemTracker(3)
ws.Range("Q7") = gemTracker(4)
ws.Range("Q8") = gemTracker(5)
ws.Range("Q9") = gemTracker(6)
ws.Range("Q10") = gemTracker(7)

'Update total gem types matched
Dim i As Integer
Dim typesMatched
typesMatched = 0
For i = 1 To 7
    ws.Range("T6") = ws.Range("T6") + gemTracker(i)
    If gemTracker(i) > 0 Then
        typesMatched = typesMatched + 1
    End If
Next i

'Set types matched score
ws.Range("U6") = typesMatched

'score gems with bonus
ws.Range("R4") = ws.Range("R4") + ws.Range("Q4") * ws.Range("N4")
ws.Range("R5") = ws.Range("R5") + ws.Range("Q5") * ws.Range("N5")
ws.Range("R6") = ws.Range("R6") + ws.Range("Q6") * ws.Range("N6")
ws.Range("R7") = ws.Range("R7") + ws.Range("Q7") * ws.Range("N7")
ws.Range("R8") = ws.Range("R8") + ws.Range("Q8") * ws.Range("N8")
ws.Range("R9") = ws.Range("R9") + ws.Range("Q9") * ws.Range("N9")
ws.Range("R10") = ws.Range("R10") + ws.Range("Q10") * ws.Range("N10")

'Update turn score
ws.Range("W6") = (ws.Range("R4") + ws.Range("R5") + ws.Range("R6") + ws.Range("R7") + _
                    ws.Range("R8") + ws.Range("R9") + ws.Range("R10")) * ws.Range("W1")

'Update main score
ws.Range("W3") = ws.Range("W3") + ws.Range("W6")


'expire gem multipliers
For i = 0 To 4
    ws.Range("O4").Offset(i, 0) = ws.Range("O4").Offset(i, 0) - 1
    If ws.Range("O4").Offset(i, 0).Value < 1 Then
        ws.Range("O4").Offset(i, -1) = 1
        'add bolding / remove bold?
    End If
Next i

'expire main multiplier
ws.Range("X1") = ws.Range("X1") - 1
If ws.Range("X1") < 1 Then
    ws.Range("W1") = 1
End If

'set gem multpliers and durations
For i = 1 To 5
    If gemTracker(i) > 3 Then
        ws.Range("O3").Offset(i, 0) = Int(gemTracker(i) / 3)
        ws.Range("O3").Offset(i, -1) = gemTracker(i) - 2
    End If
Next i

'set bonus
If typesMatched > 1 Then
    ws.Range("X1") = Int(typesMatched / 2)
    ws.Range("W1") = 1 + Int(typesMatched / 1.25)
End If

End Sub

