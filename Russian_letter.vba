Sub Main()
        Dim chletter As Char
        chletter = InputBox("Enter letter")
        Select Case chletter
            Case "�", "�", "�", "�", "�", "�", "�", "�", "�", "�"
                MsgBox("vowel")
            Case "�", "�", "�", "�", " �", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�"
                MsgBox("consonant")
            Case Else
                MsgBox("Enter russian letter")


        End Select
