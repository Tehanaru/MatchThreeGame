Attribute VB_Name = "func_updateBoard"
Public Function updateBoard(board As Range, gemTracker() As Integer) As Integer

Dim matches As Integer
matches = convertMatched(board)

Dim colGems(1 To 10) As Integer

Dim cols As Integer
Dim rows As Integer

Dim tempMatches As Integer
Dim stableGems As Integer
Dim rowCeiling As Integer
Dim matchComplete As Boolean

Dim gemsAdded As Integer
gemsAdded = 0
'Dim debugLoop As Integer

For cols = 1 To 10
    rowCeiling = 0
'    debugLoop = 0
Do

'    debugLoop = debugLoop + 1
'    If debugLoop > 50 Then
'        Exit Function
'    End If

    If board(1, cols) = "%" Then
        Do While board(rows, cols) = "%"
            rowCeiling = rowCeiling + 1
            rows = rows + 1
        Loop
    End If

    rows = 10
    stableGems = 0
    tempMatches = 0
    
    'count up to first matched gem
    Do Until isMatched(board(rows, cols)) Or board(rows, cols) = "%"
        rows = rows - 1
        stableGems = stableGems + 1
        If rows = 0 Then
            rows = 1
            Exit Do
        End If
    Loop
    
    
    'count matched gems, end on first unmatched gem
    Do While isMatched(board(rows, cols))
        Select Case board(rows, cols)
            Case "T":
                gemTracker(1) = gemTracker(1) + 1
            Case "U":
                gemTracker(2) = gemTracker(2) + 1
            Case "V":
                gemTracker(3) = gemTracker(3) + 1
            Case "W":
                gemTracker(4) = gemTracker(4) + 1
            Case "X":
                gemTracker(5) = gemTracker(5) + 1
            Case "Y":
                gemTracker(6) = gemTracker(6) + 1
            Case "Z":
                gemTracker(7) = gemTracker(7) + 1
        End Select
        rows = rows - 1
        tempMatches = tempMatches + 1
        If rows = 0 Then
            rows = 1
            Exit Do
        End If
    Loop
    
    rowCeiling = rowCeiling + tempMatches
    
    'push all gems down over matched gems
    If tempMatches > 0 Then
        For rows = rows To 1 Step -1
            board(rows + tempMatches, cols) = board(rows, cols)
        Next rows
    
        For rows = 1 To rowCeiling
            board(rows, cols) = "%"
        Next rows
    End If
    

Loop Until stableGems = 10 - rowCeiling

If rowCeiling > 0 Then
    For rows = 1 To rowCeiling
        board(rows, cols) = randLetter()
        gemsAdded = gemsAdded + 1
    Next rows
End If

Next cols

updateBoard = gemsAdded

End Function

