Attribute VB_Name = "func_matchCompare"
Public Function matchCompare(letter As String) As String

Select Case letter
    Case "A": matchCompare = "T"
    Case "B": matchCompare = "U"
    Case "C": matchCompare = "V"
    Case "D": matchCompare = "W"
    Case "E": matchCompare = "X"
    Case "F": matchCompare = "Y"
    Case "G": matchCompare = "Z"
    
    Case "T": matchCompare = "A"
    Case "U": matchCompare = "B"
    Case "V": matchCompare = "C"
    Case "W": matchCompare = "D"
    Case "X": matchCompare = "E"
    Case "Y": matchCompare = "F"
    Case "Z": matchCompare = "G"
End Select

End Function

