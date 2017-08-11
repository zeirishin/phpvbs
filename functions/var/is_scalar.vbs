Function is_scalar(str)

    is_scalar = false
    If isArray(str) or isObject(str) Then Exit Function
    if isNull(str) Then Exit Function
    is_scalar = true

End Function
