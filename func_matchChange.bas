Attribute VB_Name = "func_matchChange"
Public Function matchChange(letter As String) As String

Select Case letter
    Case "A": matchChange = "T"
    Case "B": matchChange = "U"
    Case "C": matchChange = "V"
    Case "D": matchChange = "W"
    Case "E": matchChange = "X"
    Case "F": matchChange = "Y"
    Case "G": matchChange = "Z"
    Case Else: matchChange = letter
End Select

End Function

