Attribute VB_Name = "func_randBetween"
Public Function randBetween(lowerbound As Integer, upperbound As Integer) As Integer

randBetween = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)

End Function

