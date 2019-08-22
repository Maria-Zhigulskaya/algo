        Dim IntC As Integer
        Dim StrS As String = " "
        For IntC = 1 To 100
            Console.Write(IntC)
            Console.Write(StrS)
            If IntC Mod 10 = 0 Then
                Console.WriteLine()
            End If
           Next