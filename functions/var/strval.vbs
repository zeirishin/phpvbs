Function strval(ByVal str)

    strval = false
    If isArray(str) or isObject(str) Then Exit Function
    strval = Cstr(str)

End Function
