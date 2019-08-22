Dim S, a, n As Integer
        n = 0
        Do While True
            a = InputBox(" enter number")
            If (a = 0) Then Exit Do
            S = S + a
            n = n + 1
        Loop
        MsgBox(S / n)