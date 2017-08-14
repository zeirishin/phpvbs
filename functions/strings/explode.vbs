Function explode(delimiter,string,limit)

    explode = false
    If len(delimiter) = 0 Then Exit Function
    If len(limit) = 0 Then limit = 0

    If limit > 0 Then
        explode = Split(string,delimiter,limit)
    Else
        explode = Split(string,delimiter)
    End If

End Function
