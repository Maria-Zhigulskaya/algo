Module Module1
    Public Sub MaxDate(ByVal D1 As Date, D2 As Date, D3 As Date, ByRef maxD1 As Date,ByRef maxD2 As Date)
        If D1 >= D2 Then
            maxD1 = D1
            If D2 >= D3 Then
                maxD2 = D2
            Else maxD2 = D3
            End If
        Else maxD1 = D2
            If D1 >= D3 Then
                maxD2 = D1
            Else
                maxD2 = D3
            End If
        End If
    End Sub
    Sub Main()
        Dim D1, D2, D3 As Date
        Dim maxD1, maxD2 As Date
        D1 = #3/2/2002#

        D2 = #2/5/2005#
        D3 = #8/8/2012#
        MaxDate(D1, D2, D3, maxD1, maxD2)
        Console.WriteLine(maxD1 & " " & maxD2)



    End Sub

End Module