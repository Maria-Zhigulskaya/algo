Module Module1

    Public Structure LStudent
        Public NameStudent As String
        Public MiddleName As String
        Public LastNAme As String
        Public BirthDay As Date
        Public AverageScore As Single

    End Structure
    Public Structure StudentGroup
        Public NumberGroup As UShort
        Public NameDirection As String
        Public NumberCourse As UShort
        Public ListStudents() As LStudent
    End Structure
    Public Function BestStudentsGroup(ByVal groups() As StudentGroup) As StudentGroup
        Dim C1, C2 As Integer
        Dim ListStudents() As LStudent
        Dim MaxMark, CurrentMark As Single
        Dim BestGroup As StudentGroup = groups(0)
        For C1 = 0 To UBound(groups)
            ListStudents = groups(C1).ListStudents
            For C2 = 0 To UBound(ListStudents)
                CurrentMark += ListStudents(C2).AverageScore
            Next
            If (MaxMark < CurrentMark \ (UBound(ListStudents) + 1)) Then
                MaxMark = CurrentMark \ (UBound(ListStudents) + 1)
                BestGroup = groups(C1)
            End If
            CurrentMark = 0

        Next

        Return BestGroup
    End Function