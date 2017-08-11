Function unset(ByRef val)

    If isObject(val) Then
        set val = Nothing
    Else
        val = null
    End If

End Function
