Attribute VB_Name = "sub_initBoard"
Public Sub initBoard(Optional aN As Integer = 20, Optional bN As Integer = 20, _
                    Optional cN As Integer = 20, Optional dN As Integer = 20, _
                    Optional eN As Integer = 20, Optional fN As Integer = 20, _
                    Optional gN As Integer = 20)

Dim debugLoop As Integer

Dim wb As Workbook
Set wb = ThisWorkbook

Dim ws As Worksheet
Set ws = wb.Sheets("Board")

Dim board As Range
Set board = ws.Range("A1:J10")

Dim c As Range
Dim tempLetter As String

Dim omitA As Boolean
Dim omitB As Boolean
Dim omitC As Boolean
Dim omitD As Boolean
Dim omitE As Boolean
Dim omitF As Boolean
Dim omitG As Boolean

Start:
debugLoop = debugLoop + 1
If debugLoop > 100 Then
    Exit Sub
End If

If aN = 0 Then
    omitA = True
Else
    omitA = False
End If

If bN = 0 Then
    omitB = True
Else
    omitB = False
End If

If cN = 0 Then
    omitC = True
Else
    omitC = False
End If

If dN = 0 Then
    omitD = True
Else
    omitD = False
End If

If eN = 0 Then
    omitE = True
Else
    omitE = False
End If

If fN = 0 Then
    omitF = True
Else
    omitF = False
End If
If gN = 0 Then
    omitG = True
Else
    omitG = False
End If




For Each c In board
    If aN = 0 Then
        omitA = True
    End If

    If bN = 0 Then
        omitB = True
    End If

    If cN = 0 Then
        omitC = True
    End If

    If dN = 0 Then
        omitD = True
    End If

    If eN = 0 Then
        omitE = True
    End If
    
    If fN = 0 Then
        omitF = True
    End If
    
    If gN = 0 Then
        omitG = True
    End If
    
    tempLetter = randLetter(omitA, omitB, omitC, omitD, omitE, omitF, omitG)
    
    Select Case tempLetter
        Case "A":
            c.Value = tempLetter
            aN = aN - 1
        Case "B":
            c.Value = tempLetter
            bN = bN - 1
        Case "C":
            c.Value = tempLetter
            cN = cN - 1
        Case "D":
            c.Value = tempLetter
            dN = dN - 1
        Case "E":
            c.Value = tempLetter
            eN = eN - 1
        Case "F":
            c.Value = tempLetter
            fN = fN - 1
        Case "G":
            c.Value = tempLetter
            gN = gN - 1
    End Select
Next

Dim activity As Integer

Dim gemTracker(1 To 7) As Integer

Do
    activity = updateBoard(board, gemTracker)
    debugLoop = debugLoop + 1
    If debugLoop > 100 Then
        Exit Sub
    End If
Loop Until activity = 0

If canAnyMatch(board) = False Then
    GoTo Start
End If

ws.Range("W3") = 0 ' total score


'Last Turn's gems
ws.Range("Q4") = 0
ws.Range("Q5") = 0
ws.Range("Q6") = 0
ws.Range("Q7") = 0
ws.Range("Q8") = 0
ws.Range("Q9") = 0
ws.Range("Q10") = 0

'Overall Gems Game
ws.Range("R4") = 0
ws.Range("R5") = 0
ws.Range("R6") = 0
ws.Range("R7") = 0
ws.Range("R8") = 0
ws.Range("R9") = 0
ws.Range("R10") = 0

'Score History Total Gems
ws.Range("T6") = 0
ws.Range("T7") = 0
ws.Range("T8") = 0
ws.Range("T9") = 0
ws.Range("T10") = 0

'Score History Matches
ws.Range("U6") = 0
ws.Range("U7") = 0
ws.Range("U8") = 0
ws.Range("U9") = 0
ws.Range("U10") = 0

'Score History Multipliers
ws.Range("V6") = 0
ws.Range("V7") = 0
ws.Range("V8") = 0
ws.Range("V9") = 0
ws.Range("V10") = 0

'Score History Score Gained
ws.Range("W6") = 0
ws.Range("W7") = 0
ws.Range("W8") = 0
ws.Range("W9") = 0
ws.Range("W10") = 0

'Multipliers
ws.Range("N4") = 1
ws.Range("N5") = 1
ws.Range("N6") = 1
ws.Range("N7") = 1
ws.Range("N8") = 1
ws.Range("N9") = 1
ws.Range("N10") = 1


End Sub



