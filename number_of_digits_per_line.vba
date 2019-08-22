Module Module1
    Public Function GetCountNum(ByVal StrS As String) As Short
        Dim ShrNum As Short
        Dim intC As Integer
        For intC = 1 To Len(StrS)
            If IsNumeric(GetChar(StrS, intC)) Then
                ShrNum = ShrNum + 1
            Else

            End If
        Next
        Return ShrNum
    End Function
    Sub Main()
        Dim StrX As String = "HNK419MV6DS23"
        StrX = MsgBox(GetCountNum(StrX))
    End Sub

End Module