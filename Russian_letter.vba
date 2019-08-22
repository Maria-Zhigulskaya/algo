Sub Main()
        Dim chletter As Char
        chletter = InputBox("Enter letter")
        Select Case chletter
            Case "à", "î", "ó", "å", "¸", "û", "ý", "ÿ", "è", "þ"
                MsgBox("vowel")
            Case "á", "â", "ã", "ä", " æ", "ç", "ê", "ë", "ì", "í", "ï", "ð", "ñ", "ò", "ô", "ì", "õ", "ö", "÷", "ù", "ø", "ü", "ú", "é"
                MsgBox("consonant")
            Case Else
                MsgBox("Enter russian letter")


        End Select
