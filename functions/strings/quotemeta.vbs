Function quotemeta(byVal str)

    Dim pattern : pattern = array("￥",".","+","*","?","[","^","]","$","(",")")

    Dim key
    For key = 0 to uBound(pattern)
        str = Replace(str, pattern(key),"￥" & pattern(key))
    Next
    quotemeta = str

End Function
