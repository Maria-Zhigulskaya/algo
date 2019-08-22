Module Module1

    Public Function OverWord1(ByVal strLetter As String, ByVal strBase As String) As Boolean
        Dim intC As Integer
        For intC = 1 To Len(strBase)
            If strLetter = GetChar(strBase, intC) Then
                Return True
            End If
        Next
        Return False
    End Function


    Public Function LetterCount(ByVal strLetter As String, ByVal strBase As String) As Integer
        Dim intB, Count As Integer
        For intB = 1 To Len(strBase)
            If strLetter = GetChar(strBase, intB) Then
                Count += 1
            End If
        Next
        Return Count
    End Function


    Public Function OverWord(ByVal strBase As String, ByVal strTarget As String) As Boolean
        Dim intA As Integer
        For intA = 1 To Len(strTarget)
            If OverWord1(GetChar(strTarget, intA), strBase) <> True Then
                Return False
            End If
        Next

        For intA = 1 To Len(strTarget)
            If LetterCount(GetChar(strTarget, intA), strTarget) > 1 Then
                If LetterCount(GetChar(strTarget, intA), strTarget) > LetterCount(GetChar(strTarget, intA), strBase) Then
                    Return False
                End If
            End If
        Next
        Return True
    End Function


    Sub Main()

        Dim strBase As String = "макароны"
        Dim strTarget As String = "ка"
        MsgBox(OverWord(strBase, strTarget))



    End Sub

End Module
