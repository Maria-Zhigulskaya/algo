Module Module1
    Public Function GetMax(ByVal arrNums(,) As Integer) As Integer
        Dim intC1, intC2 As Integer
        Dim intMax, intIndex As Integer
        intMax = arrNums(0, 0)
        intIndex = 0

        For intC1 = 1 To UBound(arrNums, 1)
            For intC2 = 1 To UBound(arrNums, 2)
                If arrNums(intC1, intC2) > intMax Then
                    intMax = arrNums(intC1, intC2)
                End If
            Next
        Next
        intIndex = intMax
        Return intIndex
    End Function

    Sub Main()
        Dim arrNums(2, 2) As Integer
        Dim otvet As String
        arrNums(0, 0) = 1
        arrNums(0, 1) = 2
        arrNums(0, 2) = 3
        arrNums(1, 0) = 4
        arrNums(1, 1) = 5
        arrNums(1, 2) = 6
        arrNums(2, 0) = 7
        arrNums(2, 1) = 8
        arrNums(2, 2) = 9
        otvet = GetMax(arrNums)
        MsgBox("maximum element:" & answer)

    End Sub
End Module