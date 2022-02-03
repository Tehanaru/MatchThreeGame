Attribute VB_Name = "func_isMatched"
Public Function isMatched(cell As Range) As Boolean


Select Case cell.Value
    Case "T": isMatched = True
    Case "U": isMatched = True
    Case "V": isMatched = True
    Case "W": isMatched = True
    Case "X": isMatched = True
    Case "Y": isMatched = True
    Case "Z": isMatched = True
    Case Else: isMatched = False
End Select

End Function

