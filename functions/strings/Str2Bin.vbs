Function Str2Bin(strData)

    Dim strChar,strHex
    For i = 1 To Len(strData)
        strChar = Mid(strData, i, 1)
        strHex = CStr(Hex(Asc(strChar)))
        Select Case Len(strHex)
            Case 1 '1Byte
                Str2Bin = Str2Bin & ChrB(CInt("&H" & Mid(strHex, 1, 1)))
            Case 2 '1Byte
                Str2Bin = Str2Bin & ChrB(CInt("&H" & Mid(strHex, 1, 2)))
            Case 4 '2Byte
                Str2Bin = Str2Bin & ChrB(CInt("&H" & Mid(strHex, 1, 2)))
                Str2Bin = Str2Bin & ChrB(CInt("&H" & Mid(strHex, 3, 2)))
        End Select
    Next
End Function
