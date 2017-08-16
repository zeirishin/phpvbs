Function levenshtein( str1, str2 )

    Dim s,l,t,i,j,m,n,u
    Dim a,tmp

    s = str_split(str1,1)
    u = str_split(str2,1)

    If isArray(s) Then l = count(s,") Else l = 0
    If isArray(u) Then t = count(u,") Else t = 0

    If is_empty(l) or is_empty(t) Then
        If [>](l , t) Then levenshtein = l
        If [>](t , l) Then levenshtein = t
        If isEmpty(levenshtein) Then levenshtein = 0
        Exit Function
    End If

    ReDim a(l)
    For i = 0 to l
        [] a(i),t
    Next

    For i = l to 0 Step -1
       a(i)(0) = i
    Next

    For i = t to 0 Step -1
       a(0)(i) = i
    Next

    i = 0
    m = l

    Do While(i < m)

        j = 0
        n = t

        Do While(j < n)
            tmp = a(i)(j + 1) + 1
            If tmp > a(i+1)(j) + 1 Then tmp = a(i+1)(j) + 1
            If tmp > a(i)(j) + intval([!=](s(i) ,u(j))) Then tmp = a(i)(j) + intval([!=](s(i) ,u(j)))
            a(i+1)(j+1) = tmp

            j = j + 1
        Loop

        i = i +1
    Loop

    levenshtein = a(l)(t)

End Function
