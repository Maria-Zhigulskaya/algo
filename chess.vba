Module Module1
    Public Structure Scell
        Public Column As Short
        Public Row As Short
    End Structure
    Public Function Rook(ByVal scFrom As Scell, ByVal scTo As Scell) As Boolean
        If scFrom.Column = scTo.Column Or scFrom.Row = scTo.Row Then
            Return True
        Else
            Return False
        End If
    End Function