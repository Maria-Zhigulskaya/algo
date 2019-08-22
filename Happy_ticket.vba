Dim TramNumber As Integer
        Dim A, B, C, Z, Y, X As Integer
        TramNumber = 569497
        X = TramNumber Mod 10

        Y = Fix(TramNumber / 10) Mod 10

        Z = Fix(TramNumber / 100) Mod 10
        C = Fix(TramNumber / 1000) Mod 10
        B = Fix(TramNumber / 10000) Mod 10
        A = Fix(TramNumber / 100000) Mod 10
        If (X + Y + Z) = (A + B + C) Then
            MsgBox("happy ticket")
        Else
            MsgBox("unlucky ticket")
        End If