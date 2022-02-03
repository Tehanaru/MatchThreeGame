Attribute VB_Name = "func_randLetter"
Public Function randLetter(Optional omitA As Boolean = False, _
                            Optional omitB As Boolean = False, _
                            Optional omitC As Boolean = False, _
                            Optional omitD As Boolean = False, _
                            Optional omitE As Boolean = False, _
                            Optional omitF As Boolean = False, _
                            Optional omitG As Boolean = False) As String

'Dim randInt As Integer
'randInt = randBetween(1, 5)
'
'Select Case randInt
'    Case 1: randLetter = "A"
'    Case 2: randLetter = "B"
'    Case 3: randLetter = "C"
'    Case 4: randLetter = "D"
'    Case 5: randLetter = "E"
'    Case Else: randLetter = "*"
'End Select

Dim randInt As Integer

Dim maxItems As Integer
maxItems = 7

If omitA = True Then
    maxItems = maxItems - 1
End If

If omitB = True Then
    maxItems = maxItems - 1
End If

If omitC = True Then
    maxItems = maxItems - 1
End If

If omitD = True Then
    maxItems = maxItems - 1
End If

If omitE = True Then
    maxItems = maxItems - 1
End If

If omitF = True Then
    maxItems = maxItems - 1
End If

If omitG = True Then
    maxItems = maxItems - 1
End If

If maxItems = 0 Then
    maxItems = 7
    omitA = False
    omitB = False
    omitC = False
    omitD = False
    omitE = False
    omitF = False
    omitG = False
End If

ReDim letterList(1 To maxItems) As String

Dim i As Integer
i = 1
    
If omitA = False Then
    letterList(i) = "A"
    i = i + 1
End If

If omitB = False Then
    letterList(i) = "B"
    i = i + 1
End If

If omitC = False Then
    letterList(i) = "C"
    i = i + 1
End If

If omitD = False Then
    letterList(i) = "D"
    i = i + 1
End If

If omitE = False Then
    letterList(i) = "E"
    i = i + 1
End If

If omitF = False Then
    letterList(i) = "F"
    i = i + 1
End If

If omitG = False Then
    letterList(i) = "G"
    i = i + 1
End If

randLetter = letterList(randBetween(1, maxItems))

End Function


