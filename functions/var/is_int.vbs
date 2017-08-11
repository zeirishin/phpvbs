Function is_int(str)

    is_int = false
    if Not isNumeric(str) Then Exit Function
    if str < 0 Then Exit Function
    is_int = (varType(str) = 2 or varType(str) = 3)

End Function
