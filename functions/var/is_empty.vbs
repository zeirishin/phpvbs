
Function is_empty(s)

    is_empty = false

    If isArray(s) Then
        If uBound(s) < 0 Then
            is_empty = true
            Exit Function
        Else
            Exit Function
        End If
    End If

    If isObject(s) Then
        If s.Count < 1 Then
            is_empty = true
            Exit Function
        Else
            Exit Function
        End If
    End If

    If isEmpty(s) or isNull(s) Then is_empty = true
    If s = empty Then is_empty = true

End Function
