Function lcfirst(byVal str)

    Dim tmp
    tmp = left(str,1)
    tmp = Lcase(tmp)
    lcfirst = tmp & Mid(str,2)

End Function
