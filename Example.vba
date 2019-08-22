 
        Dim IntC As Integer 
        Dim SngRes As Single
        Dim IntSign As Integer = 1
        Dim SngX As Single
        SngX = InputBox("enter number")
        For IntC = 0 To 4
     
        SngRes = SngRes + IntSign * (6 - IntC) * SngX ^ IntC
        IntSign = -IntSign
         Next
        Console.WriteLine(SngRes)
