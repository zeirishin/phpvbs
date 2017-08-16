Function stripos( haystack, needle, offset)

    Dim i
    stripos = false

    If len(offset) = 0 Then
        offset = 1
    End If

    i = inStr(offset,haystack,needle,vbTextCompare)

    If i > 0 Then
        stripos = i
    End If

End Function
