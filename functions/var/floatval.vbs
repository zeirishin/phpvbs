Function floatval(str)

    floatval = false
    If isArray(str) or isObject(str) Then Exit Function
    If not isNumeric(str) Then Exit Function
    floatval = CDbl(str)

End Function
